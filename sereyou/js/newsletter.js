document.querySelectorAll('.news form').forEach(function (form) {
  form.addEventListener('submit', function (e) {
    e.preventDefault();
    var button = form.querySelector('button');
    var input = form.querySelector('input');
    input.setAttribute('disabled', 'true');
    button.textContent = 'Iscritta ✓';
    button.setAttribute('disabled', 'true');
  });
});
