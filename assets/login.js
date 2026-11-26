(() => {
  const loginForm = document.getElementById('loginForm');
  const registerForm = document.getElementById('registerForm');
  const username = document.getElementById('username');
  const password = document.getElementById('password');
  const error = document.getElementById('authError');
  const toggle = document.getElementById('togglePassword');
  const switcher = document.getElementById('authSwitch');
  const title = document.getElementById('loginTitle');
  const intro = document.getElementById('formIntro');
  let registering = false;
  toggle.addEventListener('click', () => {
    const show = password.type === 'password';
    password.type = show ? 'text' : 'password';
    toggle.setAttribute('aria-pressed', String(show));
    toggle.setAttribute('aria-label', show ? 'Ocultar senha' : 'Mostrar senha');
    password.focus();
  });
  switcher.addEventListener('click', () => {
    registering = !registering;
    loginForm.hidden = registering;
    registerForm.hidden = !registering;
    error.hidden = true;
    title.innerHTML = registering ? 'Criar seu acesso<span>.</span>' : 'Bem-vindo de volta<span>.</span>';
    intro.textContent = registering ? 'Use o código de convite fornecido pelo administrador.' : 'Entre com suas credenciais para acessar o painel.';
    switcher.textContent = registering ? 'Já tem uma conta? Entrar' : 'Ainda não tem acesso? Criar conta com convite';
    (registering ? document.getElementById('newUsername') : username).focus();
  });
  async function submit(form, endpoint, button, payload) {
    error.hidden = true;
    if (!form.reportValidity()) return;
    button.disabled = true; button.classList.add('is-loading'); button.setAttribute('aria-busy', 'true');
    try {
      const response = await fetch(endpoint, {method:'POST',headers:{'Content-Type':'application/json','Accept':'application/json'},credentials:'same-origin',cache:'no-store',body:JSON.stringify(payload)});
      let result = {}; try { result = await response.json(); } catch (_) {}
      if (!response.ok) { error.textContent = result.error || 'Não foi possível concluir. Tente novamente.'; error.hidden = false; return; }
      window.location.replace('/');
    } catch (_) { error.textContent = 'Não foi possível conectar ao servidor. Verifique sua conexão e tente novamente.'; error.hidden = false; }
    finally { button.disabled = false; button.classList.remove('is-loading'); button.removeAttribute('aria-busy'); }
  }
  loginForm.addEventListener('submit', event => {
    event.preventDefault();
    if (!username.value.trim() || !password.value) { error.textContent='Preencha usuário e senha.'; error.hidden=false; return; }
    submit(loginForm, '/api/login', document.getElementById('submitButton'), {username:username.value.trim(),password:password.value});
  });
  registerForm.addEventListener('submit', event => {
    event.preventDefault();
    submit(registerForm, '/api/register', document.getElementById('registerButton'), {username:document.getElementById('newUsername').value.trim(),password:document.getElementById('newPassword').value,inviteCode:document.getElementById('inviteCode').value});
  });
})();
