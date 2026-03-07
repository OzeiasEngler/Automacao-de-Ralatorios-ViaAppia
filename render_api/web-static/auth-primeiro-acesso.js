/**
 * Primeiro acesso: exige troca de senha quando must_change_password.
 * Quando obrigatório, o modal não pode ser fechado (clique fora, Esc ou botão).
 *
 * Se o front for servido em outro domínio (ex.: www) e a API no Render, defina a base da API
 * para que login e demais chamadas apontem ao Render: <meta name="artesp-api-base" content="https://SUA_API.onrender.com">
 * ou antes deste script: <script>window.ARTESP_API_BASE = 'https://SUA_API.onrender.com';</script>
 */
(function() {
  var base = window.ARTESP_API_BASE || (function() {
    var m = document.querySelector('meta[name="artesp-api-base"]');
    return (m && m.getAttribute('content')) ? m.getAttribute('content').trim() : '';
  })();
  if (base) {
    base = base.replace(/\/$/, '');
    var origFetch = window.fetch;
    window.fetch = function(url, opts) {
      if (typeof url === 'string' && url.charAt(0) === '/') url = base + url;
      return origFetch.call(this, url, opts);
    };
  }

  var modal = null;
  var obrigatorio = true;
  var getToken = function() { return window.localStorage && window.localStorage.getItem('artesp_token'); };
  var getAuthHeader = function() {
    var t = getToken();
    return t ? { 'Authorization': 'Bearer ' + t } : {};
  };

  function createModal() {
    if (modal) return modal;
    var wrap = document.createElement('div');
    wrap.id = 'artesp-trocar-senha-overlay';
    wrap.style.cssText = 'position:fixed;inset:0;background:rgba(0,0,0,.75);display:flex;align-items:center;justify-content:center;z-index:99999;padding:20px;';
    wrap.addEventListener('click', function(e) {
      if (e.target === wrap && !obrigatorio) hide();
    });
    document.addEventListener('keydown', function(e) {
      if (e.key !== 'Escape' || !modal || modal.style.display !== 'flex') return;
      if (obrigatorio) e.preventDefault();
      else hide();
    });
    wrap.innerHTML =
      '<div class="artesp-trocar-senha-card" style="background:var(--bg-surface, #1a2236);border:1px solid var(--border-color, #2a3a52);border-radius:12px;padding:28px;max-width:400px;width:100%;box-shadow:0 8px 40px rgba(0,0,0,.5);">' +
      '<h3 style="margin:0 0 8px;font-size:1.1rem;color:var(--text-primary,#e8edf5);"><i class="fas fa-key"></i> Troca de senha obrigatória</h3>' +
      '<p style="margin:0 0 20px;font-size:.85rem;color:var(--text-muted,#5a6f8a);">É seu primeiro acesso. Defina uma nova senha.</p>' +
      '<form id="artesp-form-trocar-senha">' +
      '<div style="margin-bottom:14px;"><label style="display:block;font-size:.78rem;color:var(--text-secondary);margin-bottom:4px;">Senha atual (enviada para você)</label><input type="password" id="artesp-senha-atual" required placeholder="Senha atual" style="width:100%;padding:10px 14px;background:var(--bg-surface-2,#1f2b3e);border:1px solid var(--border-color);border-radius:8px;color:var(--text-primary);font-size:.9rem;"></div>' +
      '<div style="margin-bottom:14px;"><label style="display:block;font-size:.78rem;color:var(--text-secondary);margin-bottom:4px;">Nova senha (mín. 6 caracteres)</label><input type="password" id="artesp-nova-senha" required minlength="6" placeholder="Nova senha" style="width:100%;padding:10px 14px;background:var(--bg-surface-2);border:1px solid var(--border-color);border-radius:8px;color:var(--text-primary);font-size:.9rem;"></div>' +
      '<div style="margin-bottom:18px;"><label style="display:block;font-size:.78rem;color:var(--text-secondary);margin-bottom:4px;">Confirmar nova senha</label><input type="password" id="artesp-nova-senha-conf" required minlength="6" placeholder="Repita a nova senha" style="width:100%;padding:10px 14px;background:var(--bg-surface-2);border:1px solid var(--border-color);border-radius:8px;color:var(--text-primary);font-size:.9rem;"></div>' +
      '<div id="artesp-trocar-senha-feedback" style="margin-bottom:12px;font-size:.85rem;"></div>' +
      '<button type="submit" id="artesp-btn-trocar-senha" style="width:100%;padding:12px;background:var(--accent,#3b82f6);color:#fff;border:none;border-radius:8px;font-weight:600;cursor:pointer;"><i class="fas fa-check"></i> Alterar senha</button>' +
      '</form></div>';
    document.body.appendChild(wrap);
    modal = wrap;
    var card = wrap.querySelector('.artesp-trocar-senha-card');
    if (card) card.addEventListener('click', function(e) { e.stopPropagation(); });

    document.getElementById('artesp-form-trocar-senha').addEventListener('submit', function(e) {
      e.preventDefault();
      var senhaAtual = document.getElementById('artesp-senha-atual').value.trim();
      var novaSenha = document.getElementById('artesp-nova-senha').value.trim();
      var novaConf = document.getElementById('artesp-nova-senha-conf').value.trim();
      var fb = document.getElementById('artesp-trocar-senha-feedback');
      var btn = document.getElementById('artesp-btn-trocar-senha');
      fb.textContent = '';
      if (novaSenha !== novaConf) {
        fb.textContent = 'Nova senha e confirmação não conferem.';
        fb.style.color = 'var(--danger, #ef4444)';
        return;
      }
      if (novaSenha.length < 6) {
        fb.textContent = 'Nova senha deve ter no mínimo 6 caracteres.';
        fb.style.color = 'var(--danger, #ef4444)';
        return;
      }
      if (novaSenha === senhaAtual) {
        fb.textContent = 'A nova senha deve ser diferente da senha atual.';
        fb.style.color = 'var(--danger, #ef4444)';
        return;
      }
      btn.disabled = true;
      btn.innerHTML = '<i class="fas fa-spinner fa-spin"></i> Alterando...';
      btn.style.cursor = 'wait';
      fetch('/auth/trocar-senha', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json', ...getAuthHeader() },
        credentials: 'include',
        body: JSON.stringify({ senha_atual: senhaAtual, nova_senha: novaSenha })
      })
        .then(function(r) { return r.json().then(function(data) { return { ok: r.ok, data: data }; }); })
        .then(function(res) {
          if (res.ok) {
            fb.textContent = 'Senha alterada. Redirecionando...';
            fb.style.color = 'var(--success, #10b981)';
            setTimeout(function() { location.reload(); }, 800);
          } else {
            fb.textContent = (res.data && res.data.detail) || 'Erro ao alterar senha.';
            fb.style.color = 'var(--danger, #ef4444)';
            btn.disabled = false;
            btn.innerHTML = '<i class="fas fa-check"></i> Alterar senha';
            btn.style.cursor = 'pointer';
          }
        })
        .catch(function() {
          fb.textContent = 'Falha na conexão. Tente novamente.';
          fb.style.color = 'var(--danger, #ef4444)';
          btn.disabled = false;
          btn.innerHTML = '<i class="fas fa-check"></i> Alterar senha';
          btn.style.cursor = 'pointer';
        });
    });
    return modal;
  }

  function show(obrigatorioTroca) {
    obrigatorio = obrigatorioTroca !== false;
    createModal().style.display = 'flex';
  }

  function hide() {
    if (modal) modal.style.display = 'none';
  }

  function check() {
    if (!getToken()) return;
    fetch('/auth/me', { credentials: 'include', headers: getAuthHeader() })
      .then(function(r) { return r.json(); })
      .then(function(data) {
        if (data && data.must_change_password) show();
      })
      .catch(function() {});
  }

  window.ArtespTrocarSenha = { check: check, show: show, hide: hide };
})();
