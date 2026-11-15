import os
from datetime import datetime, date
from functools import wraps
from flask import Flask, render_template, request, redirect, url_for, flash, session, abort, make_response, send_file
from flask_sqlalchemy import SQLAlchemy
from werkzeug.security import generate_password_hash, check_password_hash
from flask_wtf.csrf import CSRFProtect
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment
from openpyxl.utils import get_column_letter
from apscheduler.schedulers.background import BackgroundScheduler
import sqlite3, io

app = Flask(__name__)
app.config["SECRET_KEY"] = os.environ.get("SECRET_KEY", "troque-esta-chave-em-producao")
app.config["SQLALCHEMY_DATABASE_URI"] = "sqlite:///" + os.path.join(app.instance_path, "estoque.db")
app.config["SQLALCHEMY_TRACK_MODIFICATIONS"] = False
os.makedirs(app.instance_path, exist_ok=True)
db = SQLAlchemy(app)
csrf = CSRFProtect(app)
BACKUP_DIR = os.path.join(app.instance_path, 'backups')
os.makedirs(BACKUP_DIR, exist_ok=True)

class User(db.Model):
    id = db.Column(db.Integer, primary_key=True)
    username = db.Column(db.String(80), unique=True, nullable=False)
    password_hash = db.Column(db.String(255), nullable=False)
    role = db.Column(db.String(20), default="usuario", nullable=False)
    active = db.Column(db.Boolean, default=True)
    def set_password(self, password): self.password_hash = generate_password_hash(password)
    def check_password(self, password): return check_password_hash(self.password_hash, password)

class Supplier(db.Model):
    id = db.Column(db.Integer, primary_key=True)
    name = db.Column(db.String(120), nullable=False)
    document = db.Column(db.String(30))
    phone = db.Column(db.String(40))
    email = db.Column(db.String(120))
    address = db.Column(db.String(200))

class Product(db.Model):
    id = db.Column(db.Integer, primary_key=True)
    code = db.Column(db.String(50), unique=True, nullable=False)
    name = db.Column(db.String(140), nullable=False)
    description = db.Column(db.Text)
    category = db.Column(db.String(80))
    unit = db.Column(db.String(20), default="un")
    quantity = db.Column(db.Integer, default=0, nullable=False)
    minimum = db.Column(db.Integer, default=0, nullable=False)
    cost = db.Column(db.Float, default=0)
    supplier_id = db.Column(db.Integer, db.ForeignKey("supplier.id"))
    supplier = db.relationship("Supplier")

class Movement(db.Model):
    id = db.Column(db.Integer, primary_key=True)
    product_id = db.Column(db.Integer, db.ForeignKey("product.id"), nullable=False)
    user_id = db.Column(db.Integer, db.ForeignKey("user.id"), nullable=False)
    kind = db.Column(db.String(10), nullable=False)
    quantity = db.Column(db.Integer, nullable=False)
    note = db.Column(db.String(250))
    created_at = db.Column(db.DateTime, default=datetime.now, nullable=False)
    product = db.relationship("Product")
    user = db.relationship("User")

def login_required(fn):
    @wraps(fn)
    def wrapper(*args, **kwargs):
        if "user_id" not in session: return redirect(url_for("login"))
        return fn(*args, **kwargs)
    return wrapper

def admin_required(fn):
    @wraps(fn)
    def wrapper(*args, **kwargs):
        if "user_id" not in session: return redirect(url_for("login"))
        if session.get("role") != "admin": abort(403)
        return fn(*args, **kwargs)
    return wrapper

@app.context_processor
def inject_globals():
    return {"low_stock_count": Product.query.filter(Product.quantity <= Product.minimum).count() if app.config.get("DB_READY") else 0}

@app.route("/login", methods=["GET","POST"])
def login():
    if request.method == "POST":
        user = User.query.filter_by(username=request.form.get("username","").strip(), active=True).first()
        if user and user.check_password(request.form.get("password","")):
            session.clear(); session["user_id"]=user.id; session["username"]=user.username; session["role"]=user.role
            return redirect(url_for("dashboard"))
        flash("Usuário ou senha inválidos.", "danger")
    return render_template("login.html")

@app.route("/logout")
def logout():
    session.clear(); return redirect(url_for("login"))


@app.route("/alterar-senha", methods=["GET","POST"])
@login_required
def change_password():
    user=User.query.get_or_404(session["user_id"])
    if request.method=="POST":
        current=request.form.get("current_password","")
        new=request.form.get("new_password","")
        confirm=request.form.get("confirm_password","")
        if not user.check_password(current): flash("A senha atual está incorreta.", "danger")
        elif len(new)<8: flash("A nova senha deve ter pelo menos 8 caracteres.", "danger")
        elif new != confirm: flash("A confirmação não corresponde à nova senha.", "danger")
        elif current == new: flash("A nova senha deve ser diferente da atual.", "warning")
        else:
            user.set_password(new); db.session.commit()
            flash("Senha alterada com sucesso.", "success")
            return redirect(url_for("dashboard"))
    return render_template("change_password.html")

@app.route("/admin")
@admin_required
def admin_dashboard():
    backups=sorted([f for f in os.listdir(BACKUP_DIR) if f.endswith(".db")], reverse=True)
    return render_template("admin.html", backups=backups[:10])

@app.route("/backup/agora", methods=["POST"])
@admin_required
def backup_now():
    path=create_backup()
    flash("Backup criado: "+os.path.basename(path), "success")
    return redirect(url_for("admin_dashboard"))

@app.route("/relatorios/excel")
@login_required
def export_excel():
    start=request.args.get("inicio"); end=request.args.get("fim")
    query=Movement.query
    if start:
        try: query=query.filter(Movement.created_at >= datetime.combine(date.fromisoformat(start),datetime.min.time()))
        except ValueError: pass
    if end:
        try: query=query.filter(Movement.created_at <= datetime.combine(date.fromisoformat(end),datetime.max.time()))
        except ValueError: pass
    rows=query.order_by(Movement.created_at.desc()).all()
    wb=Workbook(); ws=wb.active; ws.title="Movimentações"
    ws.append(["Data","Código","Produto","Tipo","Quantidade","Custo unitário (R$)","Valor total (R$)","Usuário","Observação"])
    for cell in ws[1]:
        cell.font=Font(bold=True,color="FFFFFF")
        cell.fill=PatternFill("solid",fgColor="263B63")
        cell.alignment=Alignment(horizontal="center")
    for m in rows:
        ws.append([m.created_at.strftime("%d/%m/%Y %H:%M"),m.product.code,m.product.name,m.kind,m.quantity,m.product.cost,m.quantity*m.product.cost,m.user.username,m.note or ""])
    for col in ws.columns:
        letter=get_column_letter(col[0].column)
        ws.column_dimensions[letter].width=min(max(max(len(str(c.value or "")) for c in col)+2,12),34)
    ws.freeze_panes="A2"; ws.auto_filter.ref=ws.dimensions
    output=io.BytesIO(); wb.save(output); output.seek(0)
    return send_file(output,as_attachment=True,download_name="relatorio_movimentacoes.xlsx",mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

@app.route("/")
@login_required
def dashboard():
    products=Product.query.order_by(Product.name).all()
    low=[p for p in products if p.quantity <= p.minimum]
    recent=Movement.query.order_by(Movement.created_at.desc()).limit(8).all()
    return render_template("dashboard.html", products=products, low=low, recent=recent)

@app.route("/produtos")
@login_required
def products():
    q=request.args.get("q","").strip()
    query=Product.query
    if q: query=query.filter(db.or_(Product.name.ilike(f"%{q}%"), Product.code.ilike(f"%{q}%")))
    return render_template("products.html", products=query.order_by(Product.name).all())

@app.route("/produtos/novo", methods=["GET","POST"])
@app.route("/produtos/<int:pid>/editar", methods=["GET","POST"])
@login_required
def product_form(pid=None):
    product=Product.query.get_or_404(pid) if pid else Product()
    if request.method=="POST":
        try:
            product.code=request.form["code"].strip()
            product.name=request.form["name"].strip()
            product.description=request.form.get("description","").strip()
            product.category=request.form.get("category","").strip()
            product.unit=request.form.get("unit","un").strip()
            product.minimum=max(0,int(request.form.get("minimum",0)))
            product.cost=max(0,float(request.form.get("cost",0) or 0))
            sid=request.form.get("supplier_id")
            product.supplier_id=int(sid) if sid else None
            if not product.id: product.quantity=0; db.session.add(product)
            db.session.commit(); flash("Produto salvo. Use Entrada para registrar o saldo inicial.", "success")
            return redirect(url_for("products"))
        except (ValueError, KeyError):
            db.session.rollback(); flash("Verifique os campos e se o código já está em uso.", "danger")
    return render_template("product_form.html", product=product, suppliers=Supplier.query.order_by(Supplier.name).all())

@app.route("/produtos/<int:pid>/excluir", methods=["POST"])
@admin_required
def product_delete(pid):
    p=Product.query.get_or_404(pid)
    if Movement.query.filter_by(product_id=p.id).first():
        flash("Não é possível excluir: o produto possui movimentações registradas.", "warning")
    else:
        db.session.delete(p); db.session.commit(); flash("Produto excluído.", "success")
    return redirect(url_for("products"))

@app.route("/movimentacoes", methods=["GET","POST"])
@login_required
def movements():
    if request.method=="POST":
        p=Product.query.get_or_404(int(request.form["product_id"]))
        kind=request.form["kind"]; qty=int(request.form["quantity"])
        if kind not in ("entrada","saida") or qty <= 0:
            flash("Tipo ou quantidade inválida.", "danger")
        elif kind=="saida" and qty > p.quantity:
            flash("Estoque insuficiente para essa saída.", "danger")
        else:
            p.quantity += qty if kind=="entrada" else -qty
            db.session.add(Movement(product=p,user_id=session["user_id"],kind=kind,quantity=qty,note=request.form.get("note","").strip()))
            db.session.commit(); flash("Movimentação registrada com sucesso.", "success")
        return redirect(url_for("movements"))
    history=Movement.query.order_by(Movement.created_at.desc()).limit(200).all()
    return render_template("movements.html", products=Product.query.order_by(Product.name).all(), history=history)

@app.route("/fornecedores", methods=["GET","POST"])
@login_required
def suppliers():
    if request.method=="POST":
        s=Supplier(name=request.form["name"].strip(),document=request.form.get("document"),phone=request.form.get("phone"),email=request.form.get("email"),address=request.form.get("address"))
        db.session.add(s); db.session.commit(); flash("Fornecedor cadastrado.", "success"); return redirect(url_for("suppliers"))
    return render_template("suppliers.html", suppliers=Supplier.query.order_by(Supplier.name).all())

@app.route("/fornecedores/<int:sid>/excluir", methods=["POST"])
@admin_required
def supplier_delete(sid):
    s=Supplier.query.get_or_404(sid)
    if Product.query.filter_by(supplier_id=s.id).first(): flash("Fornecedor vinculado a produto; altere o cadastro antes de excluir.", "warning")
    else: db.session.delete(s); db.session.commit(); flash("Fornecedor excluído.", "success")
    return redirect(url_for("suppliers"))

@app.route("/usuarios", methods=["GET","POST"])
@admin_required
def users():
    if request.method=="POST":
        username=request.form["username"].strip()
        password=request.form["password"]
        role=request.form.get("role","usuario")
        if len(password)<6: flash("A senha deve ter pelo menos 6 caracteres.", "danger")
        elif role not in ("admin","usuario"): flash("Perfil inválido.", "danger")
        elif User.query.filter_by(username=username).first(): flash("Esse usuário já existe.", "danger")
        else:
            u=User(username=username,role=role); u.set_password(password); db.session.add(u); db.session.commit(); flash("Usuário criado.", "success")
        return redirect(url_for("users"))
    return render_template("users.html", users=User.query.order_by(User.username).all())

@app.route("/usuarios/<int:uid>/excluir", methods=["POST"])
@admin_required
def user_delete(uid):
    u=User.query.get_or_404(uid)
    if u.id==session["user_id"]: flash("Você não pode excluir sua própria conta.", "warning")
    elif u.role=="admin" and User.query.filter_by(role="admin",active=True).count()<=1: flash("Não é permitido excluir o último administrador.", "warning")
    elif Movement.query.filter_by(user_id=u.id).first(): u.active=False; db.session.commit(); flash("Usuário desativado para preservar o histórico.", "success")
    else: db.session.delete(u); db.session.commit(); flash("Usuário excluído.", "success")
    return redirect(url_for("users"))

@app.route("/relatorios")
@login_required
def reports():
    start=request.args.get("inicio"); end=request.args.get("fim")
    query=Movement.query
    if start:
        try: query=query.filter(Movement.created_at >= datetime.combine(date.fromisoformat(start),datetime.min.time()))
        except ValueError: pass
    if end:
        try: query=query.filter(Movement.created_at <= datetime.combine(date.fromisoformat(end),datetime.max.time()))
        except ValueError: pass
    rows=query.order_by(Movement.created_at.desc()).all()
    total_entradas=sum(m.quantity*m.product.cost for m in rows if m.kind=="entrada")
    total_saidas=sum(m.quantity*m.product.cost for m in rows if m.kind=="saida")
    return render_template("reports.html", rows=rows, start=start or "", end=end or "", total_entradas=total_entradas, total_saidas=total_saidas)

@app.errorhandler(403)
def forbidden(e): return render_template("error.html", message="Acesso restrito ao administrador."),403

with app.app_context():
    db.create_all()
    if not User.query.filter_by(username="admin").first():
        admin=User(username="admin",role="admin"); admin.set_password("admin123"); db.session.add(admin); db.session.commit()
    app.config["DB_READY"]=True

def create_backup():
    stamp=datetime.now().strftime("%Y%m%d_%H%M%S")
    destination=os.path.join(BACKUP_DIR,f"estoque_{stamp}.db")
    source=os.path.join(app.instance_path,"estoque.db")
    with sqlite3.connect(source) as src_conn:
        with sqlite3.connect(destination) as dst_conn:
            src_conn.backup(dst_conn)
    backups=sorted([os.path.join(BACKUP_DIR,f) for f in os.listdir(BACKUP_DIR) if f.endswith(".db")],key=os.path.getmtime,reverse=True)
    for old_file in backups[30:]:
        try: os.remove(old_file)
        except OSError: pass
    return destination

def start_backup_scheduler():
    scheduler=BackgroundScheduler(daemon=True)
    scheduler.add_job(create_backup,"interval",hours=24,id="daily_backup",replace_existing=True)
    scheduler.start()
    return scheduler

if __name__=="__main__":
    scheduler=start_backup_scheduler()
    app.run(host="0.0.0.0",port=5000,debug=False,use_reloader=False)
