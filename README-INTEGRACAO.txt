ATUALIZAÇÃO DA TELA DE LOGIN — DUARTE PRODUTIVIDADE COMERCIAL

Este pacote é um patch para o projeto existente. Não substitua o restante da aplicação.

1. Substitua o app_server.py na raiz do projeto pelo app_server.py deste pacote.
2. Adicione login.html na raiz do projeto.
3. Adicione login.css e login.js dentro da pasta assets existente.
4. Mantenha assets/brand-logo.png (a marca enviada está incluída neste pacote).
5. No Railway, mantenha/configure as variáveis privadas DASHBOARD_USER e DASHBOARD_PASSWORD.
   Use credenciais fortes e exclusivas. A senha não fica gravada nos arquivos.
6. Faça commit e aguarde o novo deploy.

A tela usa cookie de sessão HttpOnly, SameSite=Strict, com validade de 8 horas. O cookie
recebe o atributo Secure quando o Railway informa HTTPS. As sessões ficam em memória e
serão encerradas quando o processo reiniciar. /api/health permanece público e sem dados
sensíveis para permitir o monitoramento do estado do serviço.

A autenticação foi validada com credenciais corretas/incorretas e acesso aos recursos
protegidos. O pacote depende dos módulos src/ e demais arquivos da aplicação já existentes.
