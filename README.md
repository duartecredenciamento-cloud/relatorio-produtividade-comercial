# Duarte · Produtividade Comercial

Painel web Python para ler matrizes `.xlsm` e exibir produtividade, credenciamento, qualidade e metas vencidas. A interface busca os dados uma vez e aplica os filtros localmente; listas extensas são paginadas.

## Execução local

1. Instale Python 3.11 ou superior.
2. No terminal, dentro desta pasta, execute `pip install -r requirements.txt`.
3. Coloque o cadastro de responsáveis em `dados/clientes_credenciamento_responsaveis.xlsx` e as matrizes em `dados/matrizes/`.
4. Execute `python app_server.py` e abra `http://localhost:8765`.

## Railway

1. Publique este projeto no repositório e faça o Railway construir pela raiz, usando o `Dockerfile`.
2. **Não inclua matrizes, cadastro de clientes nem senhas na imagem ou no Git.** Os arquivos de dados são excluídos por `.dockerignore` e `.gitignore`.
3. Crie um volume privado no Railway montado em `/data`. Dentro dele, mantenha:
   - `/data/matrizes/` com as matrizes `.xlsm`;
   - `/data/clientes_credenciamento_responsaveis.xlsx` com o cadastro de responsáveis;
   - `/data/appdata/` para persistir marcações e metas manuais.
4. Configure estas variáveis no serviço Railway:
   - `MATRICES_DIR=/data/matrizes`
   - `RESPONSIBLES_FILE=/data/clientes_credenciamento_responsaveis.xlsx`
   - `APPDATA_DIR=/data/appdata`
   - `DASHBOARD_USER` com um usuário não trivial;
   - `DASHBOARD_PASSWORD` com uma senha forte e única.
5. Faça deploy. O processo escuta em `0.0.0.0` e usa a variável `PORT` fornecida pela plataforma.

Em ambiente publicado, a autenticação HTTP Basic é obrigatória: configure as duas credenciais antes de liberar o domínio. `/api/health` permanece público e só informa o estado técnico da leitura; os dados do painel e as ações de escrita exigem autenticação. Não reutilize a senha de proteção das planilhas como senha do painel.

## Regras ajustadas

- **Inativos no credenciamento:** somente “Não se aplica”, “Cliente sem interesse” e “Aguardando editais”. O cartão de inativos aparece por último e em vermelho.
- **Em andamento:** estados preenchidos que não sejam esses três inativos nem concluídos.
- **Metas pendentes:** classificação pelo prazo da fórmula da planilha, sem depender da confirmação manual da bolinha: verde até 7 dias (oculto), laranja de 8 a 10 dias e vermelho acima de 10 dias. Datas ausentes não entram na lista.
- Os filtros da página atualizam os cartões e listas no navegador sem solicitar uma nova leitura das planilhas.

## Testes

Execute `python -m unittest discover -s tests -v` para validar as regras de status, responsáveis e prazos.
