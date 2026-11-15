# Estoque Local — versão avançada

Sistema Flask + SQLite para múltiplos usuários em rede local.

## Recursos avançados
- Alteração de senha pela interface (mínimo 8 caracteres e validação da senha atual).
- Proteção CSRF nos formulários.
- Exportação de relatórios de movimentação para Excel (.xlsx).
- Resumo financeiro estimado por período, com base no custo unitário cadastrado.
- Painel com gráficos de estoque (Chart.js).
- Backup automático do SQLite a cada 24 horas enquanto o servidor estiver ativo.
- Até 30 backups recentes e opção de backup manual na área Administração.

## Instalação no Windows
```powershell
python -m venv venv
venv\Scripts\activate
pip install -r requirements.txt
python app.py
```
Acesse `http://127.0.0.1:5000`. O banco fica em `instance/estoque.db` e os backups em `instance/backups/`.

## Primeiro acesso
Usuário: `admin` | Senha: `admin123`. Altere imediatamente em **Alterar senha**.

## Rede local
No servidor, execute `ipconfig`, localize o IPv4 e libere a porta TCP 5000 no Firewall do Windows apenas para a rede privada. Os demais computadores acessam `http://IP-DO-SERVIDOR:5000`. Mantenha o servidor ligado. Não compartilhe o arquivo SQLite diretamente por pasta de rede.

## Backups
O agendador interno cria uma cópia a cada 24 horas enquanto `app.py` estiver em execução e mantém as 30 mais recentes. Faça cópias periódicas da pasta `instance/backups` em outro disco ou dispositivo externo. Teste regularmente a restauração.

## Segurança para ambiente empresarial
Configure uma `SECRET_KEY` forte via variável de ambiente, HTTPS, servidor WSGI adequado (por exemplo Waitress no Windows), permissões restritas no diretório `instance` e backup externo. Não exponha a porta diretamente à internet.
Os valores financeiros são estimativas pelo custo cadastrado e não substituem controles fiscais, contábeis ou de faturamento.
