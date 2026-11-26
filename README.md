# Duarte · Produtividade Comercial — atualização de acesso e visual

Esta versão reúne o servidor, o painel e a tela de login/cadastro. O cadastro de usuários é protegido por código de convite; senhas são guardadas com hash PBKDF2 no volume persistente do Railway, nunca em texto simples. O usuário e a senha administrativos configurados por ambiente continuam funcionando.

## Railway

1. Substitua os arquivos do repositório pelos deste pacote e publique a atualização.
2. Mantenha o volume persistente montado em `/data`.
3. Configure as variáveis privadas já usadas pelo painel: `DASHBOARD_USER`, `DASHBOARD_PASSWORD`, `MATRICES_DIR=/data/matrizes`, `RESPONSIBLES_FILE=/data/clientes_credenciamento_responsaveis.xlsx`, `APPDATA_DIR=/data/appdata`, `GOOGLE_DRIVE_FOLDER_ID`, `GOOGLE_SERVICE_ACCOUNT_JSON` e `DRIVE_SYNC_SECONDS=900`.
4. Crie `DASHBOARD_INVITE_CODE` como um código longo e aleatório, compartilhe-o apenas com as pessoas que podem criar contas e salve as variáveis. Sem essa variável, o cadastro permanece fechado; o login administrativo continua disponível.
5. O arquivo `accounts.json` é criado em `APPDATA_DIR` no primeiro cadastro. Por isso, o volume persistente em `/data` é necessário para preservar as contas após deploy/reinício.

A conta administrativa configurada por `DASHBOARD_USER`/`DASHBOARD_PASSWORD` não depende do cadastro. Cada conta criada recebe sessão individual, expira após 8 horas e pode sair pelo botão da interface. Código de convite não substitui autorização interna: revogue/rotacione o código depois de usá-lo.

## Dados

Este pacote não inclui matrizes, cadastro de clientes, logs de operação ou credenciais. Não adicione planilhas, chaves JSON, senhas ou arquivos de dados ao Git. O código continua usando `GOOGLE_DRIVE_FOLDER_ID` e `GOOGLE_SERVICE_ACCOUNT_JSON` para sincronizar matrizes, quando configurados.

## Execução local

Instale Python 3.11+, execute `pip install -r requirements.txt` e `python app_server.py`. Para cadastro local, configure também `DASHBOARD_INVITE_CODE`.

## Testes

Execute `python -m unittest discover -s tests -v`.
