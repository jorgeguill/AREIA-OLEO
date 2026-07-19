/* ==========================================================================
   Painel do Forno — módulo plugável
   --------------------------------------------------------------------------
   Contrato de integração (para o SIGA e para uso standalone):

     const instancia = PainelForno.mount(container, opcoes);
     // ...
     instancia.unmount();   // remove o módulo do container

   - container: elemento DOM ou seletor CSS onde o módulo será montado.
   - opcoes.storagePrefix : prefixo das chaves no localStorage (default "painel-forno:").
   - opcoes.firebasePath  : prefixo dos nós no Realtime Database (default "").

   O módulo não depende de nada global do host além de (opcional) `firebase`,
   carregado pelo host quando a sincronização em tempo real for desejada.
   Todo o estado é isolado por instância e todas as buscas de DOM são feitas
   dentro do container — nenhum id/seletor vaza para fora do módulo.
   ========================================================================== */
(function (global) {
  'use strict';

  var TEMPLATE = [
    '<header>',
    '  <div class="brand">',
    '    <div class="mark"><svg viewBox="0 0 24 24" fill="none"><path d="M12 2C9 6 6 9 6 13a6 6 0 0012 0c0-2-1-3.5-2.2-5 .3 2-1 3-2.3 2C14.8 8 14 5.5 12 2z" fill="#fff" width="19"/></svg></div>',
    '    <div class="brand-text"><div class="t1">Painel Forno</div><div class="t2">Sync em tempo real</div></div>',
    '  </div>',
    '  <div style="display:flex;align-items:center;gap:10px;">',
    '    <div style="display:flex;align-items:center;font-size:11px;color:var(--muted);">',
    '      <span class="status-dot" data-status-dot></span><span data-sync-status>local</span>',
    '    </div>',
    '    <button class="icon-btn" data-btn-config>',
    '      <svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.8" width="17"><circle cx="12" cy="12" r="3"/><path d="M19.4 15a1.65 1.65 0 00.33 1.82l.06.06a2 2 0 11-2.83 2.83l-.06-.06a1.65 1.65 0 00-1.82-.33 1.65 1.65 0 00-1 1.51V21a2 2 0 01-4 0v-.09A1.65 1.65 0 009 19.4a1.65 1.65 0 00-1.82.33l-.06.06a2 2 0 11-2.83-2.83l.06-.06A1.65 1.65 0 004.6 15a1.65 1.65 0 00-1.51-1H3a2 2 0 010-4h.09A1.65 1.65 0 004.6 9a1.65 1.65 0 00-.33-1.82l-.06-.06a2 2 0 112.83-2.83l.06.06A1.65 1.65 0 009 4.6a1.65 1.65 0 001-1.51V3a2 2 0 014 0v.09a1.65 1.65 0 001 1.51 1.65 1.65 0 001.82-.33l.06-.06a2 2 0 112.83 2.83l-.06.06A1.65 1.65 0 0019.4 9a1.65 1.65 0 001.51 1H21a2 2 0 010 4h-.09a1.65 1.65 0 00-1.51 1z"/></svg>',
    '    </button>',
    '  </div>',
    '</header>',
    '<main>',
    '  <div data-alert-box></div>',
    '  <div class="hero">',
    '    <div style="font-size:11px;color:var(--muted);text-align:center;text-transform:uppercase;letter-spacing:0.08em;margin-bottom:8px;">Consumo — L/ton</div>',
    '    <div class="gauge-wrap">',
    '      <div class="gauge-track"></div>',
    '      <div class="gauge-needle" data-needle></div>',
    '      <div class="gauge-hub"></div>',
    '      <div class="gauge-center">',
    '        <div class="gauge-value mono" data-gauge-val>—</div>',
    '        <div class="gauge-unit">L/TON</div>',
    '      </div>',
    '    </div>',
    '  </div>',
    '  <section>',
    '    <div class="sec-title">Resumo</div>',
    '    <div class="card">',
    '      <div class="card-title">Areia processada</div>',
    '      <div style="display:grid;grid-template-columns:1fr 1fr;gap:10px;">',
    '        <div><div class="card-val" data-total-umida>—</div><div class="card-sub">m³ úmida</div></div>',
    '        <div><div class="card-val" data-total-seca>—</div><div class="card-sub">ton seca</div></div>',
    '      </div>',
    '    </div>',
    '  </section>',
    '  <section>',
    '    <div class="sec-title">Últimos registros</div>',
    '    <div class="card">',
    '      <div data-recent-list style="padding:0;"></div>',
    '    </div>',
    '  </section>',
    '</main>',
    '<div class="bottom-actions">',
    '  <button class="btn btn-secondary" data-btn-receb>Recebimento</button>',
    '  <button class="btn btn-primary" data-btn-turno>Registrar Turno</button>',
    '</div>',
    '<div class="overlay" data-overlay></div>',
    '<div class="sheet" data-sheet-turno>',
    '  <div class="sheet-handle"></div>',
    '  <h3>Registrar turno</h3>',
    '  <div class="sub">Areia processada, óleo e silo que recebeu.</div>',
    '  <form data-form-turno>',
    '    <div class="row2">',
    '      <div class="field"><label>Data</label><input type="date" name="data" required></div>',
    '      <div class="field"><label>Turno</label><select name="turno" required><option>Manhã</option><option>Tarde</option><option>Noite</option></select></div>',
    '    </div>',
    '    <div class="field"><label>Operador</label><input type="text" name="operador" placeholder="Nome" required></div>',
    '    <div class="subhead">Areia processada</div>',
    '    <div class="field"><label>m³ úmida</label><input type="number" step="0.1" min="0" name="m3_umida" placeholder="ex: 5.5" required></div>',
    '    <div class="subhead">Silo que recebeu</div>',
    '    <div class="field"><label>Qual silo?</label><select name="silo" required><option value="">-- Escolha --</option><option value="Silo 1 (grande)">Silo 1 (grande)</option><option value="Silo 2 (pequeno)">Silo 2 (pequeno)</option><option value="Silo 3 (pequeno)">Silo 3 (pequeno)</option><option value="Silo 4 (grande)">Silo 4 (grande)</option></select></div>',
    '    <div class="subhead">Óleo consumido</div>',
    '    <div class="row2">',
    '      <div class="field"><label>Início (L)</label><input type="number" step="1" min="0" name="tanque_ini" placeholder="850" required></div>',
    '      <div class="field"><label>Fim (L)</label><input type="number" step="1" min="0" name="tanque_fim" placeholder="620" required></div>',
    '    </div>',
    '    <div class="field"><label>Recebido (L)</label><input type="number" step="1" min="0" name="oleo_receb" value="0"></div>',
    '    <div class="field"><label>Consumo L/ton (opcional)</label><input type="number" step="0.01" min="0" name="consumo_manual" placeholder="deixa em branco para calcular automático"></div>',
    '    <div class="hint">Se deixar em branco, o app calcula automático. Se preencher, registra o valor que você digitar.</div>',
    '    <div class="sheet-actions">',
    '      <button type="button" class="btn btn-secondary" data-close>Cancelar</button>',
    '      <button type="submit" class="btn btn-primary">Salvar</button>',
    '    </div>',
    '  </form>',
    '</div>',
    '<div class="sheet" data-sheet-receb>',
    '  <div class="sheet-handle"></div>',
    '  <h3>Recebimento de óleo</h3>',
    '  <div class="sub">Registro de descarga com fornecedor.</div>',
    '  <form data-form-receb>',
    '    <div class="row2">',
    '      <div class="field"><label>Data</label><input type="date" name="data" required></div>',
    '      <div class="field"><label>Hora</label><input type="time" name="hora" required></div>',
    '    </div>',
    '    <div class="row2">',
    '      <div class="field"><label>Antes (L)</label><input type="number" step="1" min="0" name="antes" required></div>',
    '      <div class="field"><label>Depois (L)</label><input type="number" step="1" min="0" name="depois" required></div>',
    '    </div>',
    '    <div class="field"><label>Fornecedor</label><input type="text" name="fornecedor" placeholder="Nome do fornecedor" required></div>',
    '    <div class="field"><label>Recebedor</label><input type="text" name="recebedor" placeholder="Nome" required></div>',
    '    <div class="sheet-actions">',
    '      <button type="button" class="btn btn-secondary" data-close>Cancelar</button>',
    '      <button type="submit" class="btn btn-primary">Salvar</button>',
    '    </div>',
    '  </form>',
    '</div>',
    '<div class="sheet" data-sheet-config>',
    '  <div class="sheet-handle"></div>',
    '  <h3>Configurações</h3>',
    '  <div class="sub">Firebase e parâmetros da areia.</div>',
    '  <form data-form-config>',
    '    <div class="alert">⚠️ Para sincronizar dados com toda a equipe, configure o Firebase abaixo.</div>',
    '    <div class="subhead">Firebase (copie de seu projeto)</div>',
    '    <div class="field"><label>API Key</label><input type="text" name="apiKey" placeholder="AIzaSy..." value=""></div>',
    '    <div class="field"><label>Database URL</label><input type="text" name="databaseURL" placeholder="https://xxx.firebaseio.com" value=""></div>',
    '    <div class="hint">Se deixar em branco, usa localStorage (dados locais apenas).</div>',
    '    <div class="subhead">Parâmetros da areia</div>',
    '    <div class="row2">',
    '      <div class="field"><label>Umidade bruta (%)</label><input type="number" step="0.1" min="0" max="50" name="umidade" value="8" required></div>',
    '      <div class="field"><label>Densidade seca (ton/m³)</label><input type="number" step="0.01" min="0.1" name="densidade" value="1.5" required></div>',
    '    </div>',
    '    <div class="row2">',
    '      <div class="field"><label>Meta L/ton</label><input type="number" step="0.1" min="0.1" name="meta" value="8" required></div>',
    '      <div class="field"><label>Tolerância (%)</label><input type="number" step="1" min="1" name="tolerancia" value="15" required></div>',
    '    </div>',
    '    <div class="sheet-actions">',
    '      <button type="button" class="btn btn-secondary" data-close>Cancelar</button>',
    '      <button type="submit" class="btn btn-primary">Salvar</button>',
    '    </div>',
    '  </form>',
    '</div>'
  ].join('\n');

  function mount(target, options) {
    options = options || {};
    var root = typeof target === 'string' ? document.querySelector(target) : target;
    if (!root) throw new Error('PainelForno.mount: container não encontrado (' + target + ')');

    var PREFIX = options.storagePrefix || 'painel-forno:';
    var CFG = PREFIX + 'cfg';
    var TURNOS = PREFIX + 'turnos';
    var RECEB = PREFIX + 'recebimentos';
    var FB_TURNOS = (options.firebasePath || '') + 'turnos';
    var FB_RECEB = (options.firebasePath || '') + 'recebimentos';

    root.classList.add('pf-app');
    root.innerHTML = TEMPLATE;

    var $ = function (sel) { return root.querySelector(sel); };
    var $$ = function (sel) { return root.querySelectorAll(sel); };

    var state = { cfg: { umidade: 8, densidade: 1.5, meta: 8, tolerancia: 15 }, turnos: [], recebimentos: [] };
    var firebaseDb = null, useFirebase = false;

    function load(k, d) {
      try { var v = localStorage.getItem(k); return v ? JSON.parse(v) : d; } catch (e) { return d; }
    }

    function fbNode(k) {
      if (k === TURNOS) return FB_TURNOS;
      if (k === RECEB) return FB_RECEB;
      return null;
    }

    function save(k, v) {
      try { localStorage.setItem(k, JSON.stringify(v)); } catch (e) { console.error(e); }
      var node = fbNode(k);
      if (useFirebase && firebaseDb && node) {
        firebaseDb.ref(node).set(v).catch(function (e) { console.error('Firebase save error:', e); });
      }
    }

    function initFirebase() {
      var cfg = state.cfg;
      if (!cfg.apiKey || !cfg.databaseURL) { useFirebase = false; return; }
      if (typeof firebase === 'undefined') { console.warn('Firebase SDK não carregado pelo host — usando localStorage.'); useFirebase = false; return; }
      try {
        var appName = 'painel-forno-' + PREFIX;
        var existing = firebase.apps.filter(function (a) { return a.name === appName; })[0];
        var firebaseApp = existing || firebase.initializeApp(
          { apiKey: cfg.apiKey, databaseURL: cfg.databaseURL, projectId: 'painel-forno' },
          appName
        );
        firebaseDb = firebase.database(firebaseApp);
        useFirebase = true;
        syncFromFirebase();
        var st = $('[data-sync-status]'); if (st) st.textContent = 'sincronizado';
        var dot = $('[data-status-dot]'); if (dot) dot.classList.add('on');
      } catch (e) {
        console.error('Firebase init error:', e);
        useFirebase = false;
      }
    }

    function syncFromFirebase() {
      if (!useFirebase || !firebaseDb) return;
      firebaseDb.ref(FB_TURNOS).once('value').then(function (snap) {
        state.turnos = snap.val() || [];
        return firebaseDb.ref(FB_RECEB).once('value');
      }).then(function (snap) {
        state.recebimentos = snap.val() || [];
        render();
      }).catch(function (e) { console.error(e); });
    }

    function compute(t) {
      var m3u = Number(t.m3_umida), m3s = m3u * (1 - state.cfg.umidade / 100), ton = m3s * state.cfg.densidade;
      var oleo = (Number(t.tanque_ini) - Number(t.tanque_fim)) + Number(t.oleo_receb || 0);
      var lton = Number(t.consumo_manual) || null;
      if (!lton && ton > 0.01) lton = oleo / ton;
      return { m3u: m3u, m3s: m3s, ton: ton, oleo: oleo, lton: lton };
    }

    function esc(s) {
      return String(s == null ? '' : s).replace(/[&<>"']/g, function (c) {
        return { '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[c];
      });
    }

    function render() {
      var turnos = state.turnos.slice().sort(function (a, b) { return new Date(a.data) - new Date(b.data); });
      var last = turnos[turnos.length - 1], totU = 0, totT = 0;
      turnos.forEach(function (t) { totU += Number(t.m3_umida || 0); totT += Number(compute(t).ton || 0); });
      var c = last ? compute(last) : { lton: null };
      $('[data-gauge-val]').textContent = c.lton ? c.lton.toFixed(2) : '—';
      var angle = c.lton ? -135 + Math.min(270, c.lton / 20 * 270) : 0;
      $('[data-needle]').style.transform = 'translate(-50%,-100%) rotate(' + angle + 'deg)';
      $('[data-total-umida]').textContent = totU.toFixed(2);
      $('[data-total-seca]').textContent = totT.toFixed(2);
      var html = '';
      turnos.slice(-8).reverse().forEach(function (t) {
        var cc = compute(t);
        html += '<div class="entry"><div class="entry-dot turno"></div><div class="entry-main"><div class="entry-title">' + esc(t.turno) + ' • ' + esc(t.operador) + '</div><div class="entry-sub">' + esc(t.data) + ' • ' + esc(t.silo) + '</div></div><div class="entry-val">' + (cc.lton ? cc.lton.toFixed(2) : '—') + '</div></div>';
      });
      state.recebimentos.slice(-5).reverse().forEach(function (r) {
        html += '<div class="entry"><div class="entry-dot receb"></div><div class="entry-main"><div class="entry-title">Receb • ' + esc(r.fornecedor) + '</div><div class="entry-sub">' + esc(r.data) + ' ' + esc(r.hora) + '</div></div><div class="entry-val">+' + (Number(r.depois) - Number(r.antes)).toFixed(0) + 'L</div></div>';
      });
      $('[data-recent-list]').innerHTML = html || '<div style="padding:8px;color:var(--muted);font-size:12px;">Sem registros</div>';
    }

    function openSheet(sheet) {
      sheet.classList.add('open');
      $('[data-overlay]').classList.add('open');
    }

    function close() {
      $$('.sheet').forEach(function (s) { s.classList.remove('open'); });
      $('[data-overlay]').classList.remove('open');
    }

    function init() {
      state.cfg = load(CFG, state.cfg);
      state.turnos = load(TURNOS, []);
      state.recebimentos = load(RECEB, []);
      initFirebase();
      render();
    }

    // ---- wiring (todos os listeners escopados no container) ------------------
    var handlers = [];
    function on(el, ev, fn) { el.addEventListener(ev, fn); handlers.push([el, ev, fn]); }

    on($('[data-btn-turno]'), 'click', function () {
      $('[data-form-turno] [name=data]').value = new Date().toISOString().slice(0, 10);
      openSheet($('[data-sheet-turno]'));
    });

    on($('[data-btn-receb]'), 'click', function () {
      $('[data-form-receb] [name=data]').value = new Date().toISOString().slice(0, 10);
      $('[data-form-receb] [name=hora]').value = new Date().toTimeString().slice(0, 5);
      openSheet($('[data-sheet-receb]'));
    });

    on($('[data-btn-config]'), 'click', function () {
      var f = $('[data-form-config]');
      f.apiKey.value = state.cfg.apiKey || '';
      f.databaseURL.value = state.cfg.databaseURL || '';
      f.umidade.value = state.cfg.umidade;
      f.densidade.value = state.cfg.densidade;
      f.meta.value = state.cfg.meta;
      f.tolerancia.value = state.cfg.tolerancia;
      openSheet($('[data-sheet-config]'));
    });

    on($('[data-overlay]'), 'click', close);
    $$('[data-close]').forEach(function (b) { on(b, 'click', close); });

    on($('[data-form-turno]'), 'submit', function (e) {
      e.preventDefault();
      var f = e.target;
      state.turnos.push({
        data: f.data.value, turno: f.turno.value, operador: f.operador.value,
        m3_umida: Number(f.m3_umida.value), silo: f.silo.value,
        tanque_ini: Number(f.tanque_ini.value), tanque_fim: Number(f.tanque_fim.value),
        oleo_receb: Number(f.oleo_receb.value || 0),
        consumo_manual: f.consumo_manual.value ? Number(f.consumo_manual.value) : null
      });
      save(TURNOS, state.turnos);
      f.reset(); close(); render();
      alert('✅ Turno salvo!');
    });

    on($('[data-form-receb]'), 'submit', function (e) {
      e.preventDefault();
      var f = e.target;
      state.recebimentos.push({
        data: f.data.value, hora: f.hora.value,
        antes: Number(f.antes.value), depois: Number(f.depois.value),
        fornecedor: f.fornecedor.value, recebedor: f.recebedor.value
      });
      save(RECEB, state.recebimentos);
      f.reset(); close(); render();
      alert('✅ Recebimento salvo!');
    });

    on($('[data-form-config]'), 'submit', function (e) {
      e.preventDefault();
      var f = e.target;
      state.cfg = {
        apiKey: f.apiKey.value, databaseURL: f.databaseURL.value,
        umidade: Number(f.umidade.value), densidade: Number(f.densidade.value),
        meta: Number(f.meta.value), tolerancia: Number(f.tolerancia.value)
      };
      save(CFG, state.cfg);
      close();
      if (f.apiKey.value && f.databaseURL.value) { initFirebase(); }
      render();
      alert('✅ Configurações salvas!');
    });

    init();

    return {
      unmount: function () {
        handlers.forEach(function (h) { h[0].removeEventListener(h[1], h[2]); });
        handlers = [];
        root.innerHTML = '';
        root.classList.remove('pf-app');
      },
      refresh: render,
      state: state
    };
  }

  global.PainelForno = { mount: mount, version: '1.0.0-teste' };
})(typeof window !== 'undefined' ? window : this);
