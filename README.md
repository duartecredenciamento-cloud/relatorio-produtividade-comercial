# Duarte · Produtividade Comercial

Painel web para matrizes `.xlsm`, com tela de login própria, sessão encerrável, filtros locais e atualização de fontes. **O pacote não contém dados reais, planilhas de clientes, logs ou credenciais.**

## Rodar no Windows

1. Instale Python 3.11+ e execute `instalar.bat` uma vez.
2. Coloque o cadastro de responsáveis em `dados/clientes_credenciamento_responsaveis.xlsx` (ou informe outro caminho ao iniciar).
3. Execute `iniciar.ps1` no PowerShell. Ele usa por padrão a pasta oficial `G:/Meu Drive/OPERAÇÃO/1 - COMERCIAL/DEMONSTRATIVOSDIARIOS_COM/NOVA MATRIZ CREDENCIAMENTO`, solicita usuário e senha sem salvá-los no projeto e abre o painel local.
4. Acesse `http://127.0.0.1:8765`. Para encerrar, use **Sair** na aplicação ou Ctrl+C no terminal.

A unidade Google Drive precisa estar montada no PC. O botão Atualizar agora força uma nova leitura; o painel preserva o último conjunto válido quando uma fonte falha.

## Publicar no Railway com sincronização automática do Google Drive

O Railway não pode ler a unidade `G:`. Para sincronizar sem depender do computador, use a Google Drive API e um volume persistente privado.

1. Em um projeto Google Cloud aprovado pela empresa, habilite a Google Drive API e crie uma conta de serviço.
2. Compartilhe com ela, como **Leitor**, a pasta oficial que contém as matrizes `.xlsm` diretamente (sem subpastas). Copie o ID da pasta da URL.
3. Crie no Railway um volume privado montado em `/data`.
4. Configure no serviço Railway as variáveis privadas:
   - `MATRICES_DIR=/data/matrizes`
   - `RESPONSIBLES_FILE=/data/clientes_credenciamento_responsaveis.xlsx`
   - `APPDATA_DIR=/data/appdata`
   - `GOOGLE_DRIVE_FOLDER_ID=<ID da pasta>`
   - `GOOGLE_SERVICE_ACCOUNT_JSON=<conteúdo integral da chave JSON>`
   - `DRIVE_SYNC_SECONDS=900`
   - `DASHBOARD_USER=<usuário forte>`
   - `DASHBOARD_PASSWORD=<senha forte e exclusiva>`
5. Coloque **uma vez** o cadastro de responsáveis em `/data/clientes_credenciamento_responsaveis.xlsx`. O Railway sincroniza as matrizes ao iniciar, reconsulta por padrão a cada 15 minutos e baixa somente arquivos novos/alterados; **Atualizar agora** força uma verificação imediata.
6. Faça deploy e teste antes de divulgar o endereço. O estado da sincronização e os avisos aparecem no painel/log.

A integração usa escopo de leitura do Drive. A chave da conta de serviço é segredo: nunca a coloque no GitHub, no ZIP distribuído ou em mensagens. O funcionamento depende de a política do Google Workspace permitir o compartilhamento com a conta de serviço.

## Login e segurança

- A raiz redireciona para a tela própria de login quando não há uma sessão válida.
- O cookie é `HttpOnly`, `SameSite=Strict`, expira em 8 horas e recebe `Secure` em HTTPS. O logout invalida a sessão no servidor e limpa o cookie.
- Os dados, arquivos do painel e ações de escrita exigem autenticação. `/api/health` expõe apenas o estado técnico mínimo.
- Senhas e dados de cliente não ficam no código. Em Railway, configure credenciais nas variáveis privadas. Em Windows, o inicializador pede as credenciais a cada execução.

## Regras e filtros

- Inativos do credenciamento: “Não se aplica”, “Cliente sem interesse” e “Aguardando editais”; o cartão de inativos fica por último e destacado em vermelho.
- Metas pendentes: só prazos vencidos; laranja entre 8–10 dias, vermelho acima de 10; até 7 dias fica fora da lista. A classe vem da data, não da cor de formatação condicional.
- Filtros de responsável, prioridade, categoria, serviço, operadora, data, período e busca combinam com filtros visuais em cartões, barras, status e tabelas; **Limpar filtros** remove o recorte.
- Listas são paginadas e a interface não precisa reler as planilhas para aplicar filtros.

## Testes

Execute `python -m unittest discover -s tests -v`. As verificações cobrem regras de prazo/status e fluxo de login/logout.
