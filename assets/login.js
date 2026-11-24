(() => {
  const form = document.getElementById('loginForm');
  const username = document.getElementById('username');
  const password = document.getElementById('password');
  const error = document.getElementById('loginError');
  const submit = document.getElementById('submitButton');
  const toggle = document.getElementById('togglePassword');

  toggle.addEventListener('click', () => {
    const show = password.type === 'password';
    password.type = show ? 'text' : 'password';
    toggle.setAttribute('aria-pressed', String(show));
    toggle.setAttribute('aria-label', show ? 'Ocultar senha' : 'Mostrar senha');
    password.focus();
  });

  form.addEventListener('submit', async (event) => {
    event.preventDefault();
    error.hidden = true;
    const userValue = username.value.trim();
    const passwordValue = password.value;
    if (!userValue || !passwordValue) {
      error.textContent = 'Preencha o usuário e a senha para continuar.';
      error.hidden = false;
      (!userValue ? username : password).focus();
      return;
    }

    submit.disabled = true;
    submit.classList.add('is-loading');
    submit.setAttribute('aria-busy', 'true');
    try {
      const response = await fetch('/api/login', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json', 'Accept': 'application/json' },
        credentials: 'same-origin',
        cache: 'no-store',
        body: JSON.stringify({ username: userValue, password: passwordValue })
      });
      let result = {};
      try { result = await response.json(); } catch (_) { /* resposta vazia */ }
      if (!response.ok) {
        error.textContent = response.status === 401
          ? 'Usuário ou senha incorretos. Confira os dados e tente novamente.'
          : (result.error || 'Não foi possível iniciar a sessão. Tente novamente em instantes.');
        error.hidden = false;
        password.value = '';
        password.focus();
        return;
      }
      window.location.replace('/');
    } catch (_) {
      error.textContent = 'Não foi possível conectar ao servidor. Verifique sua conexão e tente novamente.';
      error.hidden = false;
    } finally {
      submit.disabled = false;
      submit.classList.remove('is-loading');
      submit.removeAttribute('aria-busy');
    }
  });
})();
