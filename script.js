/* ToolRent: живой калькулятор аренды и проверка доступности.
   Формулы совпадают с calc_quote() в server.py (целые копейки, округление «от половины вверх»).
   Браузер только показывает результат мгновенно; цены, скидки и итоги окончательно считает Flask. */
(function () {
  'use strict';
  document.documentElement.classList.add('js');

  var cfgEl = document.getElementById('rent-config');
  var CFG = cfgEl ? JSON.parse(cfgEl.textContent) : null;
  var csrfMeta = document.querySelector('meta[name="csrf-token"]');
  var CSRF = csrfMeta ? csrfMeta.content : '';
  var DAY = 86400000;

  /* ---------- вспомогательное ---------- */
  function toInt(v, d) { var n = parseInt(v, 10); return isFinite(n) ? n : (d || 0); }
  function pad(n) { return (n < 10 ? '0' : '') + n; }
  function msOf(iso) {
    var m = /^(\d{4})-(\d{2})-(\d{2})$/.exec(iso || '');
    return m ? Date.UTC(+m[1], +m[2] - 1, +m[3]) : null;
  }
  function isoOf(ms) {
    var d = new Date(ms);
    return d.getUTCFullYear() + '-' + pad(d.getUTCMonth() + 1) + '-' + pad(d.getUTCDate());
  }
  function debounce(fn, ms) { var t; return function () { var a = arguments; clearTimeout(t); t = setTimeout(function () { fn.apply(null, a); }, ms); }; }
  function fmtK(k) {
    k = Math.round(k || 0);
    var neg = k < 0; if (neg) k = -k;
    var rub = Math.floor(k / 100), kop = k % 100;
    var s = String(rub).replace(/\B(?=(\d{3})+(?!\d))/g, '\u00a0');
    if (kop) s += ',' + (kop < 10 ? '0' : '') + kop;
    return (neg ? '−' : '') + s + '\u00a0₽';
  }
  function fmtPct(bp) {
    var v = bp / 100, s = (Math.round(v * 100) / 100).toString().replace('.', ',');
    return s + '%';
  }
  function plural(n, a, b, c) { var m = n % 100, d = n % 10; return (m > 10 && m < 20) ? c : d === 1 ? a : (d >= 2 && d <= 4) ? b : c; }

  /* ---------- расчёт (зеркало calc_quote из server.py) ---------- */
  function offK(amount, bp) { return Math.floor((amount * bp + 5000) / 10000); }
  function durBp(days) {
    var bp = 0;
    (CFG.tiers || []).forEach(function (t) { if (days >= t[0]) bp = t[1]; });
    return bp;
  }
  function quote(lines, coupon, points, balance, enforceMax) {
    if (enforceMax === undefined) enforceMax = true;
    var units = 0; lines.forEach(function (l) { units += l.qty; });
    var qbp = units >= CFG.qty_min_units ? CFG.qty_bp : 0, cbp = CFG.corp_bp;
    var tot = { base: 0, dur: 0, qty: 0, corp: 0, rent: 0, deposit: 0 }, out = [];
    lines.forEach(function (l) {
      var base = l.price_k * l.qty * l.days, dbp = durBp(l.days), dur = offK(base, dbp), a = base - dur;
      var qd = offK(a, qbp), b = a - qd, cp = offK(b, cbp), rent = b - cp, dep = l.deposit_k * l.qty;
      out.push({ base: base, dur_bp: dbp, dur: dur, qty_disc: qd, corp: cp, rent: rent, deposit: dep });
      tot.base += base; tot.dur += dur; tot.qty += qd; tot.corp += cp; tot.rent += rent; tot.deposit += dep;
    });
    var auto = tot.rent, ck = 0, cok = true, cwhy = '';
    if (coupon) {
      if (auto < (coupon.min_k || 0)) { cok = false; cwhy = 'сумма аренды меньше минимальной для купона'; }
      else if (coupon.type === 'percent') ck = offK(auto, coupon.bp);
      else ck = Math.min(coupon.amount_k, auto);
    }
    var after = auto - ck, rate = CFG.rate_k;
    var maxPts = Math.floor(Math.floor(after * CFG.points_max_bp / 10000) / rate), cap;
    if (enforceMax) cap = (balance === null || balance === undefined) ? maxPts : Math.min(maxPts, balance);
    else cap = (balance === null || balance === undefined) ? Math.floor(after / rate) : Math.min(Math.floor(after / rate), balance);
    var used = Math.max(0, Math.min(toInt(points, 0), cap)), pk = used * rate, rentTotal = after - pk;
    return { lines: out, units: units, qty_bp: qbp, corp_bp: cbp, base: tot.base, dur: tot.dur, qty: tot.qty, corp: tot.corp,
             auto_rent: auto, coupon: ck, coupon_ok: cok, coupon_why: cwhy, max_points: cap, points: used, points_k: pk,
             rent: rentTotal, deposit: tot.deposit, grand: rentTotal + tot.deposit, savings: tot.base - rentTotal };
  }
  window.ToolRentCalc = { quote: quote, fmtK: fmtK, cfg: CFG };   // для проверки и отладки

  /* ---------- сервер ---------- */
  function postJSON(url, body) {
    return fetch(url, { method: 'POST', credentials: 'same-origin',
      headers: { 'Content-Type': 'application/json', 'X-CSRF-Token': CSRF }, body: JSON.stringify(body) })
      .then(function (r) { if (!r.ok) throw new Error('HTTP ' + r.status); return r.json(); });
  }

  /* ---------- одна позиция: количество, дата начала, срок, дата окончания ---------- */
  function f(el, name) { return el.querySelector('[data-f="' + name + '"]'); }
  function o(el, name) { return el.querySelector('[data-o="' + name + '"]'); }
  function setText(el, name, text) { var n = o(el, name); if (n) n.textContent = text; }

  function readLine(el) {
    var qty = toInt(f(el, 'qty').value, 0), sIso = f(el, 'start').value, eIso = f(el, 'end').value;
    var s = msOf(sIso), e = msOf(eIso), days = (s !== null && e !== null) ? Math.round((e - s) / DAY) + 1 : 0, err = '';
    var min = el.dataset.today ? msOf(el.dataset.today) : null;
    if (qty < 1) err = 'Количество должно быть не меньше 1.';
    else if (s === null || e === null) err = 'Укажите даты начала и окончания.';
    else if (days < 1) err = 'Дата окончания раньше даты начала.';
    else if (days > CFG.max_days) err = 'Максимальный срок аренды — ' + CFG.max_days + ' дн.';
    else if (min !== null && s < min) err = 'Дата начала не может быть в прошлом.';
    return { qty: qty, start: sIso, end: eIso, days: days, err: err,
             price_k: +el.dataset.priceK, deposit_k: +el.dataset.depositK };
  }

  function bindPeriod(el, onChange) {
    var qty = f(el, 'qty'), start = f(el, 'start'), days = f(el, 'days'), end = f(el, 'end');
    function endFromDays() {
      var s = msOf(start.value), d = Math.max(1, toInt(days.value, 1));
      if (s !== null) end.value = isoOf(s + (d - 1) * DAY);
    }
    function daysFromEnd() {
      var s = msOf(start.value), e = msOf(end.value);
      if (s !== null && e !== null) {
        if (e < s) { end.value = start.value; days.value = 1; } else days.value = Math.round((e - s) / DAY) + 1;
      }
    }
    function markPresets() {
      el.querySelectorAll('[data-set-days]').forEach(function (b) { b.classList.toggle('on', +b.dataset.setDays === toInt(days.value, 0)); });
    }
    qty.addEventListener('input', function () { onChange('qty'); });
    start.addEventListener('input', function () { endFromDays(); markPresets(); onChange('start'); });
    days.addEventListener('input', function () { if (toInt(days.value, 0) >= 1) { endFromDays(); markPresets(); onChange('days'); } });
    end.addEventListener('input', function () { daysFromEnd(); markPresets(); onChange('end'); });
    el.querySelectorAll('[data-set-days]').forEach(function (b) {
      b.addEventListener('click', function () { days.value = b.dataset.setDays; endFromDays(); markPresets(); onChange('days'); });
    });
    var fix = el.querySelector('[data-fix-qty]');
    if (fix) fix.addEventListener('click', function () { qty.value = fix.dataset.fixQty; onChange('qty'); });
    markPresets();
  }

  function showAvail(el, a, line) {
    var box = o(el, 'avail'); if (!box) return;
    var fix = el.querySelector('[data-fix-qty]'), qty = f(el, 'qty');
    box.className = 'avail ' + (a.err ? 'bad' : a.ok ? 'ok' : 'bad');
    if (a.err) { box.textContent = a.err; }
    else if (a.ok) { box.textContent = 'Доступно ' + a.free + ' ' + plural(a.free, 'экземпляр', 'экземпляра', 'экземпляров') + ' на этот период — выбранное количество можно взять.'; }
    else { box.textContent = a.message || 'Выбранное количество недоступно.'; }
    if (!a.err && a.free >= 0) qty.setAttribute('max', Math.max(a.free, 1));
    if (fix) { fix.hidden = !(a.free > 0 && line.qty > a.free && !a.err); fix.dataset.fixQty = a.free; fix.textContent = 'Выбрать ' + a.free + ' шт.'; }
  }

  /* ---------- страница инструмента ---------- */
  function initTool(form) {
    var el = form.querySelector('.rent-line'), btn = form.querySelector('[data-add]'), seq = 0;
    var cartUnits = toInt(el.dataset.cartUnits, 0), hint = form.querySelector('[data-o="qty-hint"]');
    function render(line, avail) {
      if (!line.err) {
        var q = quote([{ price_k: line.price_k, deposit_k: line.deposit_k, qty: line.qty, days: line.days }], null, 0, 0);
        var L = q.lines[0];
        setText(el, 'days-label', line.days + ' ' + plural(line.days, 'сутки', 'суток', 'суток'));
        setText(el, 'base', fmtK(L.base)); setText(el, 'rent', fmtK(L.rent)); setText(el, 'deposit', fmtK(L.deposit));
        setText(el, 'total', fmtK(L.rent + L.deposit)); setText(el, 'dur', '−' + fmtK(L.dur + L.corp));
        var r = el.querySelector('[data-row="dur"]'); if (r) r.hidden = !(L.dur + L.corp);
        if (hint) hint.hidden = !(cartUnits + line.qty >= CFG.qty_min_units);
      }
      if (btn) btn.disabled = !!line.err || (avail ? !avail.ok : false);
    }
    var check = debounce(function () {
      var line = readLine(el); if (line.err) { showAvail(el, { err: line.err }, line); return; }
      var my = ++seq;
      postJSON('/api/availability', { tool_id: +el.dataset.toolId, qty: line.qty, start: line.start, end: line.end })
        .then(function (a) { if (my !== seq) return; showAvail(el, a.error ? { err: a.error } : a, line); render(line, a.error ? { ok: false } : a); })
        .catch(function () { if (my === seq) { var b = o(el, 'avail'); if (b) { b.className = 'avail'; b.textContent = 'Не удалось проверить доступность. Проверка будет выполнена при добавлении.'; } } });
    }, 250);
    bindPeriod(el, function () { var l = readLine(el); render(l, null); if (btn) btn.disabled = true; check(); });
    render(readLine(el), null); if (btn) btn.disabled = true; check();
    form.addEventListener('submit', function (e) { if (readLine(el).err) e.preventDefault(); });
  }

  /* ---------- корзина ---------- */
  function initCart(form) {
    var items = [].slice.call(form.querySelectorAll('.rent-line')), seq = 0;
    var balance = toInt(form.dataset.balance, 0), couponSel = form.querySelector('[data-f="coupon"]');
    var pointsIn = form.querySelector('[data-f="points"]'), btnOk = form.querySelector('[data-checkout]');
    var state = { avail: {} };

    function coupon() {
      if (!couponSel || !couponSel.value || couponSel.value === '0') return null;
      var opt = couponSel.options[couponSel.selectedIndex];
      return { type: opt.dataset.type, bp: +opt.dataset.bp, amount_k: +opt.dataset.amountK, min_k: +opt.dataset.minK };
    }
    function setS(name, k, hideZero) {
      var n = form.querySelector('[data-s="' + name + '"]'); if (n) n.textContent = fmtK(k);
      var r = form.querySelector('[data-srow="' + name + '"]'); if (r && hideZero) r.hidden = !k;
    }
    function renderLine(el, L, line, a) {
      setText(el, 'days-label', line.days + ' ' + plural(line.days, 'сутки', 'суток', 'суток'));
      setText(el, 'base', fmtK(L.base)); setText(el, 'rent', fmtK(L.rent)); setText(el, 'deposit', fmtK(L.deposit));
      var disc = L.dur + L.qty_disc + L.corp; setText(el, 'disc', '−' + fmtK(disc));
      var r = el.querySelector('[data-row="disc"]'); if (r) r.hidden = !disc;
      var dl = o(el, 'disc-note'); if (dl) { var parts = [];
        if (L.dur) parts.push('за срок ' + fmtPct(L.dur_bp)); if (L.qty_disc) parts.push('за количество'); if (L.corp) parts.push('для бизнеса');
        dl.textContent = parts.join(', '); }
      el.classList.toggle('bad', !!(a && !a.ok) || !!line.err);
    }
    function renderTotals(Q, hasProblem) {
      setS('base', Q.base); setS('dur', -Q.dur, true); setS('qty', -Q.qty, true); setS('corp', -Q.corp, true);
      setS('coupon', -Q.coupon, true); setS('points', -Q.points_k, true);
      setS('rent', Q.rent); setS('deposit', Q.deposit); setS('grand', Q.grand); setS('savings', Q.savings, true);
      var u = form.querySelector('[data-s="units"]'); if (u) u.textContent = Q.units;
      var qn = form.querySelector('[data-o="qty-note"]'); if (qn) qn.textContent = Q.qty_bp ? ('Скидка за количество ' + fmtPct(Q.qty_bp) + ' применена: в заявке ' + Q.units + ' ед.') : ('Скидка ' + fmtPct(CFG.qty_bp) + ' за количество действует от ' + CFG.qty_min_units + ' ед. в заявке; сейчас ' + Q.units + '.');
      var cn = form.querySelector('[data-o="coupon-note"]'); if (cn) cn.textContent = Q.coupon_ok ? '' : ('Купон не применён: ' + Q.coupon_why + '.');
      var pm = form.querySelector('[data-o="points-max"]'); if (pm) pm.textContent = Q.max_points;
      if (pointsIn) { pointsIn.max = Q.max_points; if (toInt(pointsIn.value, 0) > Q.max_points) pointsIn.value = Q.max_points; }
      if (btnOk) btnOk.disabled = hasProblem || !items.length;
    }
    function recompute() {
      var lines = items.map(function (el) { return readLine(el); }), ok = lines.filter(function (l) { return !l.err; });
      var Q = quote(ok.map(function (l) { return { price_k: l.price_k, deposit_k: l.deposit_k, qty: l.qty, days: l.days }; }),
                    coupon(), pointsIn ? toInt(pointsIn.value, 0) : 0, balance);
      var j = 0, problem = false;
      items.forEach(function (el, i) {
        var l = lines[i], a = state.avail[el.dataset.k];
        if (l.err) { showAvail(el, { err: l.err }, l); problem = true; return; }
        renderLine(el, Q.lines[j++], l, a);
        if (a) { showAvail(el, a, l); if (!a.ok) problem = true; } else problem = true;   // пока нет ответа сервера, оформлять нельзя
      });
      renderTotals(Q, problem);
    }
    var sync = debounce(function () {
      var my = ++seq;
      postJSON('/api/cart/quote', { lines: items.map(function (el) { var l = readLine(el); return { k: el.dataset.k, qty: l.qty, start: l.start, end: l.end }; }),
                                    coupon: couponSel ? toInt(couponSel.value, 0) : 0, points: pointsIn ? toInt(pointsIn.value, 0) : 0 })
        .then(function (r) {
          if (my !== seq) return;
          var problem = r.has_problem;
          r.lines.forEach(function (sl) {
            var el = form.querySelector('.rent-line[data-k="' + sl.k + '"]'); if (!el) return;
            state.avail[sl.k] = { ok: sl.ok, free: sl.free, message: sl.message };
            var line = readLine(el);
            if (sl.error) { showAvail(el, { err: sl.error }, line); problem = true; return; }
            showAvail(el, state.avail[sl.k], line);
            renderLine(el, sl, line, state.avail[sl.k]);
          });
          // окончательные суммы — из Flask (если вдруг отличаются от локального расчёта, верны эти)
          if (couponSel) { [].forEach.call(couponSel.options, function (op) { var c = r.coupons.filter(function (x) { return String(x.uc) === op.value; })[0]; if (c) op.disabled = !c.ok; }); }
          renderTotals(r.quote, problem || r.lines.some(function (x) { return x.error; }));
        })
        .catch(function () { /* без ответа сервера кнопка оформления остаётся заблокированной */ });
    }, 250);
    function changed() { recompute(); sync(); }
    items.forEach(function (el) { bindPeriod(el, changed); });
    if (couponSel) couponSel.addEventListener('change', changed);
    if (pointsIn) pointsIn.addEventListener('input', changed);
    var maxBtn = form.querySelector('[data-points-max]');
    if (maxBtn) maxBtn.addEventListener('click', function () { pointsIn.value = 0; recompute(); pointsIn.value = form.querySelector('[data-o="points-max"]').textContent; changed(); });
    form.addEventListener('submit', function (e) {
      var submitter = e.submitter;
      if (submitter && submitter.hasAttribute('data-checkout') && submitter.disabled) e.preventDefault();
    });
    items.forEach(function (el) {
      if (el.dataset.availOk !== undefined) state.avail[el.dataset.k] = { ok: el.dataset.availOk === '1', free: toInt(el.dataset.free, 0), message: el.dataset.availMsg || '' };
    });
    recompute(); sync();
  }

  /* редактор характеристик в админке: добавить строку */
  document.addEventListener('click', function (e) {
    var b = e.target.closest('[data-spec-add]'); if (!b) return;
    var box = document.querySelector('[data-spec-rows]'), tpl = document.getElementById('spec-row-tpl');
    if (!box || !tpl) return;
    var node = tpl.content.firstElementChild.cloneNode(true); box.appendChild(node);
    var f = node.querySelector('input'); if (f) f.focus();
  });
  document.addEventListener('click', function (e) {
    var b = e.target.closest('[data-spec-del]'); if (!b) return;
    var row = b.closest('.spec-edit-row'); if (!row) return;
    [].forEach.call(row.querySelectorAll('input'), function (i) { i.value = ''; });
    row.hidden = true;
  });

  document.addEventListener('DOMContentLoaded', function () {
    if (!CFG) return;
    var tool = document.querySelector('form[data-tool-form]'); if (tool) initTool(tool);
    var cart = document.getElementById('cart-form'); if (cart) initCart(cart);
  });
})();



/* Искры при оплате: вспышка из кнопки, закручиваются, разлетаются и гаснут */
(function () {
  var DELAY = 1500,                                   // мс до отправки формы, чтобы эффект успел сыграть
      still = matchMedia('(prefers-reduced-motion: reduce)').matches;

  function burst(x, y, w, h) {
    var S = 3, N = 36, cv = document.createElement('canvas'), cx = cv.getContext('2d'),
        W = innerWidth, H = innerHeight, ps = [], last = performance.now();
    cv.width = Math.ceil(W / S); cv.height = Math.ceil(H / S); cx.setTransform(1 / S, 0, 0, 1 / S, 0, 0);   // рисуем «крупными пикселями»
    cv.style.cssText = 'position:fixed;inset:0;width:100%;height:100%;pointer-events:none;z-index:9999;image-rendering:pixelated';
    document.body.appendChild(cv);
    for (var i = 0; i < N; i++) {
      var an = (i + Math.random()) / N * 6.2832, sp = 220 + Math.random() * 540, big = Math.random() < .25;  // углы равномерно → не кучкуются
      ps.push({ x: x + (Math.random() - .5) * w, y: y + (Math.random() - .5) * h,
        vx: Math.cos(an) * sp, vy: Math.sin(an) * sp,
        w: (Math.random() < .5 ? -1 : 1) * (1 + Math.random() * 2.5),            // закручивание
        big: big, r: big ? 5 + Math.random() * 4 : 1.5 + Math.random() * 2.5,
        life: 1.3 + Math.random() * 1.3, age: 0, tw: 8 + Math.random() * 14, ph: Math.random() * 6.28, h: 40 + Math.random() * 14 });
    }
    (function step(t) {
      var dt = Math.min((t - last) / 1000, .05), alive = 0; last = t;
      cx.clearRect(0, 0, W, H); cx.globalCompositeOperation = 'lighter';
      ps.forEach(function (p) {
        p.age += dt; if (p.age >= p.life) return; alive++;
        var a = p.w * dt, c = Math.cos(a), s = Math.sin(a), d = Math.exp(-1.2 * dt),
            vx = (p.vx * c - p.vy * s) * d, vy = (p.vx * s + p.vy * c) * d - 22 * dt;
        p.vx = vx; p.vy = vy; p.x += vx * dt; p.y += vy * dt;
        var k = 1 - p.age / p.life, al = k * k * (.65 + .35 * Math.sin(p.age * p.tw + p.ph)),
            R = p.r * (p.big ? 4 : 3), g = cx.createRadialGradient(p.x, p.y, 0, p.x, p.y, R), ga = al * (p.big ? .8 : .5);
        g.addColorStop(0, 'hsla(' + p.h + ',100%,70%,' + ga + ')');
        g.addColorStop(.4, 'hsla(' + p.h + ',100%,55%,' + ga * .3 + ')');
        g.addColorStop(1, 'hsla(' + p.h + ',100%,50%,0)');
        cx.fillStyle = g; cx.beginPath(); cx.arc(p.x, p.y, R, 0, 6.2832); cx.fill();
        var q = Math.max(S, Math.round(p.r * 1.2 / S) * S);                      // чёткий пиксельный центр
        cx.fillStyle = 'hsla(' + (p.h + 6) + ',100%,86%,' + Math.min(1, al * 1.3) + ')';
        cx.fillRect(Math.round(p.x / S) * S - q / 2, Math.round(p.y / S) * S - q / 2, q, q);
      });
      if (alive) requestAnimationFrame(step); else cv.remove();
    })(last);
  }

  document.addEventListener('submit', function (e) {
    var f = e.target, st = f.dataset && f.dataset.fx, sb = e.submitter;
    if (!f.matches || !f.matches('form[action$="/pay"]') || st === '2' || still) return;
    e.preventDefault();
    if (st) return;                                    // повторный клик, пока играет эффект
    f.dataset.fx = '1';
    var r = (sb || f).getBoundingClientRect();
    burst(r.left + r.width / 2, r.top + r.height / 2, r.width, r.height);
    setTimeout(function () {
      f.dataset.fx = '2';
      if (f.requestSubmit) f.requestSubmit(sb || undefined); else f.submit();
    }, DELAY);
  }, true);
})();