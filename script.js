const form = document.querySelector('.login-form');
const email = document.querySelector('#email');
const password = document.querySelector('#password');
const passwordToggle = document.querySelector('.password-toggle');
const status = document.querySelector('.form-status');

passwordToggle.addEventListener('click', () => {
  const isVisible = password.type === 'text';
  password.type = isVisible ? 'password' : 'text';
  passwordToggle.textContent = isVisible ? 'Show' : 'Hide';
  passwordToggle.setAttribute('aria-label', isVisible ? 'Show password' : 'Hide password');
  passwordToggle.setAttribute('aria-pressed', String(!isVisible));
});

function setError(input, message) {
  const field = input.closest('.field');
  field.querySelector('.field__error').textContent = message;
  input.setAttribute('aria-invalid', 'true');
}

function clearError(input) {
  const field = input.closest('.field');
  field.querySelector('.field__error').textContent = '';
  input.removeAttribute('aria-invalid');
}

[email, password].forEach((input) => {
  input.addEventListener('input', () => {
    clearError(input);
    status.textContent = '';
  });
});

form.addEventListener('submit', (event) => {
  event.preventDefault();
  let isValid = true;
  const emailValue = email.value.trim();

  if (!emailValue || !email.validity.valid) {
    setError(email, 'Enter a valid email address.');
    isValid = false;
  }

  if (password.value.length < 6) {
    setError(password, 'Password must be at least 6 characters.');
    isValid = false;
  }

  if (!isValid) {
    status.textContent = '';
    (email.getAttribute('aria-invalid') ? email : password).focus();
    return;
  }

  window.location.href = 'upload.html';
});
