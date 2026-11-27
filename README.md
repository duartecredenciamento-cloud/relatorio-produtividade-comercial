# Duarte · Relatório de Produtividade Comercial

Aplicação web em Python com painel HTML/CSS/JavaScript, abas operacionais, filtros, metas manuais, revisão de marcações e sincronização opcional de matrizes `.xlsm` do Google Drive.

## Endereços e autenticação

- `/` encaminha para `/login.html` sem sessão válida e para `/painelgeral.html` com sessão válida.
- `/login.html` é a tela de acesso e cadastro por convite.
- `/painelgeral.html` e as APIs que entregam dados do painel exigem sessão válida no servidor.
- Login e cadastro válidos levam a `/painelgeral.html`. O botão **Sair** invalida a sessão no servidor.
- As sessões usam token aleatório em cookie `HttpOnly`, `SameSite=Strict` e `Secure` quando executado com `PORT` (Railway) ou quando HTTPS é indicado pelo proxy. A validade é de 8 horas.
- O login administrativo vem exclusivamente de `DASHBOARD_USER` e `DASHBOARD_PASSWORD`. Contas criadas armazenam apenas hash PBKDF2 com salt aleatório no `APPDATA_DIR`.
- O cadastro só aceita convite fornecido por `DASHBOARD_INVITE_CODE`; sem essa variável, fica fechado. O código não é enviado à interface pelo servidor.

## Configuração no Railway

Monte um volume persistente em `/data` e configure as variáveis privadas no serviço. As variáveis aceitas são:

- `PORT`: porta fornecida pelo Railway; o servidor escuta em `0.0.0.0` quando essa variável existe.
- `REQUIRE_AUTH`: a autenticação deste pacote é sempre obrigatória no servidor; não defina esta variável como forma de desativar a proteção.
- `DASHBOARD_USER` e `DASHBOARD_PASSWORD`: credenciais administrativas fortes, definidas por você.
- `DASHBOARD_INVITE_CODE`: código longo e aleatório para autorizar novos cadastros; opcional, mas obrigatório para abrir o cadastro.
- `MATRICES_DIR`: pasta persistente das matrizes, por exemplo `/data/matrizes`.
- `APPDATA_DIR`: pasta persistente de contas, estado da integração, metas e revisões, por exemplo `/data/appdata`.
- `RESPONSIBLES_FILE`: caminho do cadastro oficial de responsáveis, por exemplo `/data/clientes_credenciamento_responsaveis.xlsx`; mantenha o arquivo no armazenamento privado e não no Git.
- `GOOGLE_DRIVE_FOLDER_ID`: ID da pasta oficial do Drive **NOVA MATRIZ CREDENCIAMENTO**.
- `GOOGLE_SERVICE_ACCOUNT_JSON`: conteúdo JSON da conta de serviço fornecido como variável privada, nunca como arquivo no repositório.
- `DRIVE_SYNC_SECONDS`: intervalo mínimo de sincronização do Drive em segundos; o padrão implementado é 900 segundos quando não definido (mínimo aceito: 60).

Compartilhe a pasta oficial do Drive com a conta de serviço com permissão de leitura. Se as duas variáveis do Drive não estiverem configuradas, a aplicação usa apenas os arquivos já disponíveis no diretório de matrizes. O ciclo de leitura local mantém o intervalo padrão existente de 60 segundos e usa cache para não reabrir matrizes inalteradas.

O volume persistente preserva `accounts.json`, `metas_manuais.csv`, `marcacoes_manuais.csv` e o estado da sincronização. Nenhum valor de segredo vem preenchido neste pacote.

## Execução local

Instale Python 3.11 ou superior, execute `pip install -r requirements.txt` e inicie com `python app_server.py`. Para o primeiro acesso local, configure `DASHBOARD_USER` e `DASHBOARD_PASSWORD` no ambiente. O script `iniciar.ps1` também solicita credenciais e caminhos locais sem embutir senhas.

## Regras de prazo mantidas

Foi preservada a regra existente no projeto: até 7 dias fica dentro do prazo e fora das metas pendentes; de 8 a 10 dias é laranja; acima de 10 dias é vermelho. O campo de bolinha não confirma o prazo automaticamente e pode continuar em revisão manual. As fórmulas das matrizes examinadas tratam vencimentos de contrato/reajuste em 60 dias; não foi encontrada nelas uma fórmula específica de meta pendente, então não se substituiu a regra já implementada.

## Testes

Com as dependências instaladas, execute `python -m unittest discover -s tests -v`. Os testes usam dados sintéticos e não precisam de matrizes reais, credenciais Google ou publicação.

## Proteção de dados

Não inclua planilhas reais, arquivos `.xlsm`, `.xlsx`, `.xls`, logs, `.env`, senhas, contas de serviço ou caches no Git ou no pacote. O `.gitignore` e o `.dockerignore` excluem essas categorias; confira o estado do repositório antes de qualquer commit. Esta entrega não publica no Railway nem envia alterações ao GitHub.
