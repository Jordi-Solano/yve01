/* yve-tema v1 - el color personalizado del usuario, en todas las paginas.
 *
 * El dashboard guarda tres cosas en localStorage: yve_accent, yve_bg y
 * yve_hl_all (el modo "acentuar todo"). Hasta ahora solo el dashboard las
 * leia, asi que administracion, aprobaciones y conciliacion se quedaban con
 * el azul de fabrica: elegias carmesi y al abrir un panel aparte seguia todo
 * azul.
 *
 * Este fichero hace lo mismo que _applyCustomColors del dashboard, para
 * cualquier pagina que lo incluya. Va en el <head> a proposito: si se cargase
 * al final se veria un parpadeo azul antes de pintarse del color bueno.
 *
 * Tambien inyecta las reglas de "acentuar todo", para que una burbuja se
 * comporte igual este en la pagina que este.
 */
(function () {
  if (window.__yveTema) return;
  window.__yveTema = true;

  function hx(v) {
    v = Math.max(0, Math.min(255, Math.round(v)));
    return ('0' + v.toString(16)).slice(-2);
  }
  function haciaBlanco(v, t) { return Math.round(v + (255 - v) * t); }
  function oscurecer(v, t) { return Math.max(0, Math.round(v * (1 - t))); }
  function comp(hexa) {
    var h = hexa.replace('#', '');
    return [parseInt(h.substr(0, 2), 16), parseInt(h.substr(2, 2), 16), parseInt(h.substr(4, 2), 16)];
  }

  function aplicar() {
    var r = document.documentElement;
    var acc, bg, hl;
    try {
      acc = localStorage.getItem('yve_accent');
      bg = localStorage.getItem('yve_bg');
      hl = localStorage.getItem('yve_hl_all') === '1';
    } catch (e) { return; }

    if (bg && /^#[0-9a-f]{6}$/i.test(bg)) {
      var b = comp(bg);
      r.style.setProperty('--bg', bg);
      r.style.setProperty('--bg-r', String(b[0]));
      r.style.setProperty('--bg-g', String(b[1]));
      r.style.setProperty('--bg-b', String(b[2]));
    }
    if (acc && /^#[0-9a-f]{6}$/i.test(acc)) {
      var a = comp(acc);
      // Mismas mezclas que el dashboard: acc2 un 25% hacia el blanco, acc3 un
      // 50%, acc-dark un 20% mas oscuro. Si se cambian aqui y alli no, las dos
      // pantallas dejan de ser el mismo producto.
      r.style.setProperty('--acc', acc);
      r.style.setProperty('--acc2', '#' + hx(haciaBlanco(a[0], .25)) + hx(haciaBlanco(a[1], .25)) + hx(haciaBlanco(a[2], .25)));
      r.style.setProperty('--acc3', '#' + hx(haciaBlanco(a[0], .5)) + hx(haciaBlanco(a[1], .5)) + hx(haciaBlanco(a[2], .5)));
      r.style.setProperty('--acc-dark', '#' + hx(oscurecer(a[0], .2)) + hx(oscurecer(a[1], .2)) + hx(oscurecer(a[2], .2)));
      r.style.setProperty('--acc-r', String(a[0]));
      r.style.setProperty('--acc-g', String(a[1]));
      r.style.setProperty('--acc-b', String(a[2]));
    }

    function marcar() {
      if (!document.body) return;
      document.body.classList[hl ? 'add' : 'remove']('acentuar-todo');
    }
    if (document.body) marcar();
    else document.addEventListener('DOMContentLoaded', marcar);
  }

  // Las reglas de "acentuar todo". El dashboard ya tiene las suyas con las
  // mismas declaraciones; repetirlas no cambia nada alli y hace que en las
  // demas paginas una burbuja se comporte igual.
  function reglas() {
    if (document.getElementById('yve-tema-css')) return;
    var s = document.createElement('style');
    s.id = 'yve-tema-css';
    s.textContent =
      'body.acentuar-todo .sc,body.acentuar-todo .card,body.acentuar-todo .kc,' +
      'body.acentuar-todo .fb-kpi-card{' +
      'border-color:rgba(var(--acc-r,59),var(--acc-g,130),var(--acc-b,246),.4)!important;' +
      'background:rgba(var(--acc-r,59),var(--acc-g,130),var(--acc-b,246),.06)!important}' +
      'body.acentuar-todo .sc .sc-val,body.acentuar-todo .kv,' +
      'body.acentuar-todo .fb-kpi-val{color:var(--acc2)!important}' +
      'body.acentuar-todo .card-title,body.acentuar-todo .ct,body.acentuar-todo .kl,' +
      'body.acentuar-todo .fb-kpi-lbl{color:var(--acc2)!important;opacity:.85}';
    (document.head || document.documentElement).appendChild(s);
  }

  aplicar();
  reglas();

  // Si el usuario cambia el color en otra pestaña, esta se entera.
  window.addEventListener('storage', function (e) {
    if (e && e.key && e.key.indexOf('yve_') === 0) aplicar();
  });
})();
