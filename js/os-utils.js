(function() {
  'use strict';

  const U = {};

  U.escapeHtml = function(value) {
    return String(value == null ? '' : value).replace(/[&<>"']/g, ch => ({
      '&': '&amp;',
      '<': '&lt;',
      '>': '&gt;',
      '"': '&quot;',
      "'": '&#39;'
    }[ch]));
  };

  U.normalizeText = function(value) {
    return String(value == null ? '' : value)
      .normalize('NFD')
      .replace(/[\u0300-\u036f]/g, '')
      .toLowerCase()
      .trim();
  };

  U.normalizePlate = function(value) {
    return String(value == null ? '' : value).toUpperCase().replace(/[^A-Z0-9]/g, '');
  };

  U.parseNumberBR = function(value) {
    if (typeof value === 'number' && Number.isFinite(value)) return value;
    let s = String(value == null ? '' : value).trim();
    if (!s) return 0;
    s = s.replace(/\s/g, '').replace(/R\$/gi, '').replace(/%/g, '');
    const hasComma = s.includes(',');
    const hasDot = s.includes('.');
    if (hasComma && hasDot) {
      s = s.replace(/\./g, '').replace(',', '.');
    } else if (hasComma) {
      s = s.replace(',', '.');
    }
    s = s.replace(/[^0-9.-]/g, '');
    const n = parseFloat(s);
    return Number.isFinite(n) ? n : 0;
  };

  U.formatInputMoney = function(value) {
    const n = U.parseNumberBR(value);
    return n ? n.toFixed(2).replace('.', ',') : '0,00';
  };

  U.moeda = function(value) {
    return new Intl.NumberFormat('pt-BR', { style: 'currency', currency: 'BRL' }).format(U.parseNumberBR(value));
  };

  U.getCliente = function(os, clientes, fallbackCliente) {
    if (fallbackCliente) return fallbackCliente;
    return (clientes || []).find(c => c.id === os?.clienteId) || null;
  };

  U.getValorHoraCliente = function(cliente, fallback) {
    return U.parseNumberBR(cliente?.govValorHora || cliente?.valorHora || fallback || 0);
  };

  U.getDescontosCliente = function(cliente, os) {
    const descMO = os?.descMO != null ? U.parseNumberBR(os.descMO) : U.parseNumberBR(cliente?.govDescMO || 0);
    const descPeca = os?.descPeca != null ? U.parseNumberBR(os.descPeca) : U.parseNumberBR(cliente?.govDescPeca || 0);
    return { descMO, descPeca };
  };

  U.getVehicle = function(os, veiculos) {
    return (veiculos || []).find(v => v.id === os?.veiculoId) || {};
  };

  U.buildBudgetItems = function(os, cliente) {
    const descontos = U.getDescontosCliente(cliente, os);
    const servicos = (os?.servicos || []).map((s, index) => {
      const valorUnit = U.parseNumberBR(s.valor);
      const qtd = 1;
      const bruto = +(valorUnit * qtd).toFixed(2);
      const final = +(bruto * (1 - descontos.descMO)).toFixed(2);
      return {
        key: 'servico-' + index,
        tipo: 'servico',
        labelTipo: 'Servico',
        index,
        codigo: s.codigoTabela || s.codigo || '',
        sistema: s.sistemaTabela || s.sistema || '',
        desc: s.desc || '',
        tempo: U.parseNumberBR(s.tempo),
        qtd,
        valorUnit,
        valorBruto: bruto,
        valorFinal: final
      };
    });
    const pecas = (os?.pecas || []).map((p, index) => {
      const qtd = U.parseNumberBR(p.qtd || p.q || 1) || 1;
      const valorUnit = U.parseNumberBR(p.venda || p.valor || p.v);
      const bruto = +(qtd * valorUnit).toFixed(2);
      const final = +(bruto * (1 - descontos.descPeca)).toFixed(2);
      return {
        key: 'peca-' + index,
        tipo: 'peca',
        labelTipo: 'Peca',
        index,
        codigo: p.codigo || p.cod || '',
        sistema: p.sistemaTabela || p.sistema || '',
        desc: p.desc || p.descricao || '',
        tempo: 0,
        qtd,
        valorUnit,
        valorBruto: bruto,
        valorFinal: final
      };
    });
    return servicos.concat(pecas).filter(it => it.desc || it.codigo || it.valorBruto > 0);
  };

  U.getApprovedKeys = function(os) {
    const keys = new Set();
    const fromApproval = os?.aprovacao?.itens || os?.itensAprovados || [];
    fromApproval.forEach(item => {
      if (typeof item === 'string') keys.add(item);
      else if (item?.key) keys.add(item.key);
    });
    return keys;
  };

  U.hasApproval = function(os) {
    return !!((os?.aprovacao && Array.isArray(os.aprovacao.itens)) || Array.isArray(os?.itensAprovados));
  };

  U.splitCiliaTokens = function(textOrTokens) {
    if (Array.isArray(textOrTokens)) return textOrTokens.map(t => String(t || '').trim()).filter(Boolean);
    return String(textOrTokens || '')
      .replace(/<[^>]+>/g, ' ')
      .split(/\s+/)
      .map(t => t.trim())
      .filter(Boolean);
  };

  function isCodigoMarker(t) {
    return /^c.?d[:.]?$/i.test(U.normalizeText(t).replace(/\s/g, ''));
  }

  function isMoneyToken(t) {
    return /^-?\d{1,3}(?:\.\d{3})*,\d{2}$/.test(String(t || '')) || /^-?\d+,\d{2}$/.test(String(t || ''));
  }

  function extractCiliaPrices(text) {
    const s = String(text || '');
    const money = Array.from(s.matchAll(/R\$\s*([\d.]+,\d{2}|\d+,\d{2}|\d+\.\d{2})/gi)).map(m => m[1]);
    const desconto = s.match(/%\s*([\d.,]+)/);
    return {
      bruto: U.parseNumberBR(money[0] || 0),
      liquido: U.parseNumberBR(money.length > 1 ? money[money.length - 1] : 0),
      desconto: U.parseNumberBR(desconto?.[1] || 0)
    };
  }

  function isNumericToken(t) {
    return /^\d+(?:[,.]\d+)?$/.test(String(t || ''));
  }

  function isCiliaSummaryOrServiceBoundary(line) {
    const n = U.normalizeText(line);
    return n.includes('total pecas') ||
      n.includes('total geral') ||
      n.includes('subtotal') ||
      n.includes('mao de obra do orcamento') ||
      n.includes('total mao de obra') ||
      n.startsWith('servicos') ||
      n.includes(' servicos ');
  }

  function ciliaPieceLinesOnly(lines) {
    let inParts = false;
    let sawHeader = false;
    const selected = [];
    (lines || []).forEach(line => {
      const n = U.normalizeText(line);
      const isPartsHeader =
        n.includes('pecas e mao de obra') ||
        (n.includes('operacoes') && n.includes('descricao/codigo')) ||
        (n.includes('qtd') && n.includes('descricao/codigo') && n.includes('preco'));
      if (isPartsHeader) {
        inParts = true;
        sawHeader = true;
        return;
      }
      if (inParts && isCiliaSummaryOrServiceBoundary(line)) {
        inParts = false;
        return;
      }
      if (inParts) selected.push(line);
    });
    return sawHeader ? selected : (lines || []).filter(line => !isCiliaSummaryOrServiceBoundary(line));
  }

  function peelCiliaDescriptionAndQty(text) {
    let tokens = String(text || '').replace(/\s+/g, ' ').trim().split(' ').filter(Boolean);
    while (tokens.length && /^(T|R|P|R&I|RI)$/i.test(tokens[0])) tokens.shift();

    const leadingNumbers = [];
    while (tokens.length && isNumericToken(tokens[0])) leadingNumbers.push(tokens.shift());

    const trailingNumbers = [];
    while (tokens.length && isNumericToken(tokens[tokens.length - 1])) trailingNumbers.unshift(tokens.pop());

    let qtd = 1;
    if (leadingNumbers.length) qtd = U.parseNumberBR(leadingNumbers[leadingNumbers.length - 1]) || 1;
    if (trailingNumbers.length) qtd = U.parseNumberBR(trailingNumbers[trailingNumbers.length - 1]) || qtd || 1;

    const desc = tokens
      .join(' ')
      .replace(/\b(Oficina|Seguradora|Fornecedor|Cliente)\b/gi, ' ')
      .replace(/\s+/g, ' ')
      .trim();
    return { desc, qtd };
  }

  U.normalizeCiliaPiece = function(piece) {
    const p = { ...(piece || {}) };
    let desc = String(p.desc || p.descricao || '').replace(/\s+/g, ' ').trim();
    let codigo = String(p.codigo || p.cod || '').trim();
    const prices = extractCiliaPrices(desc);
    const codInDesc = desc.match(/C.?d[:.]\s*([A-Z0-9./-]+)/i);
    if (!codigo && codInDesc) codigo = codInDesc[1].trim();

    desc = desc
      .replace(/C.?d[:.]\s*[A-Z0-9./-]+/gi, ' ')
      .replace(/R\$\s*[\d.]+,\d{2}/gi, ' ')
      .replace(/R\$\s*\d+\.\d{2}/gi, ' ')
      .replace(/%\s*[\d.,]+/g, ' ')
      .replace(/\b(Oficina|Seguradora|Fornecedor|Cliente)\b/gi, ' ')
      .replace(/\s+/g, ' ')
      .trim();

    const peeled = peelCiliaDescriptionAndQty(desc);
    const normalized = {
      ...p,
      codigo: codigo || 'sem oem',
      desc: peeled.desc,
      qtd: U.parseNumberBR(p.qtd || p.quantidade || 0) || peeled.qtd || 1,
      venda: U.parseNumberBR(p.venda || p.valor || p.precoBruto || 0) || prices.bruto || U.parseNumberBR(p.ciliaValorLiquido || 0) || prices.liquido || 0,
      ciliaValorLiquido: U.parseNumberBR(p.ciliaValorLiquido || p.valorLiquido || 0) || prices.liquido || 0,
      ciliaDesconto: U.parseNumberBR(p.ciliaDesconto || p.desconto || 0) || prices.desconto || 0
    };
    return normalized;
  };

  U.normalizeCiliaPieces = function(pieces) {
    return (pieces || [])
      .map(U.normalizeCiliaPiece)
      .filter(p => {
        const n = U.normalizeText(p.desc);
        if (!p.desc && p.codigo === 'sem oem') return false;
        if (isCiliaSummaryOrServiceBoundary(p.desc)) return false;
        if (n === 'total' || n === 'preco' || n === 'desconto') return false;
        return p.codigo || p.desc || U.parseNumberBR(p.venda) > 0;
      });
  };

  U.parseCiliaPiecesFromLines = function(lines) {
    const out = [];
    ciliaPieceLinesOnly(lines).forEach(line => {
      const original = String(line || '').replace(/\s+/g, ' ').trim();
      if (!original || !/R\$/i.test(original)) return;
      if (isCiliaSummaryOrServiceBoundary(original)) return;
      const prices = extractCiliaPrices(original);
      if (!prices.bruto && !prices.liquido) return;

      const beforePrice = original.split(/R\$/i)[0].replace(/\s+/g, ' ').trim();
      let codigo = '';
      let descPart = beforePrice;

      const codMarker = beforePrice.match(/^(.*)\s+C.?d[:.]\s*([A-Z0-9./-]+)(?:\s+\w+)?\s*$/i);
      if (codMarker) {
        descPart = codMarker[1].trim();
        codigo = codMarker[2].trim();
      } else {
        const first = beforePrice.match(/^([A-Z0-9][A-Z0-9./-]{3,})\s+(.+)$/);
        if (first && /\d/.test(first[1]) && !/^\d+[,.]\d+$/.test(first[1]) && !/^(TOTAL|PRECO|VALOR)$/i.test(first[1])) {
          codigo = first[1].trim();
          descPart = first[2].trim();
        }
      }

      if (!codigo && !descPart) return;
      const peeled = peelCiliaDescriptionAndQty(descPart);
      if (!peeled.desc && !codigo) return;
      out.push({
        codigo,
        desc: peeled.desc,
        qtd: peeled.qtd || 1,
        venda: prices.bruto || prices.liquido || 0,
        ciliaValorLiquido: prices.liquido || 0,
        ciliaDesconto: prices.desconto || 0
      });
    });
    return U.normalizeCiliaPieces(out);
  };

  U.parseCiliaPiecesFromTokens = function(textOrTokens) {
    const tokens = U.splitCiliaTokens(textOrTokens);
    const pieces = [];
    if (!tokens.length) return pieces;

    const startIdx = tokens.findIndex((t, i) =>
      U.normalizeText(t).startsWith('operac') &&
      U.normalizeText(tokens.slice(i, i + 12).join(' ')).includes('descricao/codigo')
    );
    const totalIdx = tokens.findIndex((t, i) =>
      U.normalizeText(t) === 'total' && U.normalizeText(tokens[i + 1] || '').startsWith('pec')
    );
    const windowTokens = tokens.slice(startIdx >= 0 ? startIdx : 0, totalIdx > 0 ? totalIdx : tokens.length);
    const codeIdxs = [];
    windowTokens.forEach((t, i) => {
      if (isCodigoMarker(t) && windowTokens[i + 1]) codeIdxs.push(i);
    });

    if (codeIdxs.length) {
      const firstCodeIdx = codeIdxs[0];
      const numericBefore = [];
      for (let i = 0; i < firstCodeIdx; i++) {
        if (/^\d+(?:\.\d+)?$/.test(windowTokens[i])) numericBefore.push({ idx: i, value: windowTokens[i] });
      }
      const qtyTokens = numericBefore.slice(-codeIdxs.length);
      const firstDescIdx = qtyTokens.length ? qtyTokens[qtyTokens.length - 1].idx + 1 : Math.max(0, firstCodeIdx - 1);
      const moneyPairs = [];
      for (let i = codeIdxs[codeIdxs.length - 1] + 2; i < windowTokens.length - 1; i++) {
        if (/^R\$/i.test(windowTokens[i]) && isMoneyToken(windowTokens[i + 1])) {
          moneyPairs.push(windowTokens[i + 1]);
          i++;
        }
      }

      for (let i = 0; i < codeIdxs.length; i++) {
        const markerIdx = codeIdxs[i];
        const nextMarkerIdx = codeIdxs[i + 1] || windowTokens.length;
        const descStart = i === 0 ? firstDescIdx : codeIdxs[i - 1] + 2;
        const descTokens = windowTokens.slice(descStart, markerIdx);
        const cleanDesc = descTokens
          .filter(t => !/^(T|R|P|R&I|Oficina)$/i.test(t))
          .join(' ')
          .replace(/\s+/g, ' ')
          .trim();
        const qtd = U.parseNumberBR(qtyTokens[i]?.value || 1) || 1;
        const bruto = U.parseNumberBR(moneyPairs[i * 2] || 0);
        const liquido = U.parseNumberBR(moneyPairs[i * 2 + 1] || 0);
        if (cleanDesc || windowTokens[markerIdx + 1]) {
          pieces.push({
            codigo: windowTokens[markerIdx + 1] || '',
            desc: cleanDesc,
            qtd,
            venda: bruto || liquido,
            ciliaValorLiquido: liquido
          });
        }
        if (nextMarkerIdx <= markerIdx) break;
      }
    }

    if (pieces.length) return U.normalizeCiliaPieces(pieces);

    const text = tokens.join(' ');
    const lineRegex = /(?:[TRP](?:\s+R&I)?)?\s*[\d,.]+\s+([\d,.]+)\s+(.+?)\s+C.?d[:.]\s*([A-Z0-9./-]+)\s+\w+\s+R\$\s*([\d.,]+)\s+%\s*[\d.,]+\s+R\$\s*([\d.,]+)/gi;
    let m;
    while ((m = lineRegex.exec(text))) {
      pieces.push({
        codigo: m[3].trim(),
        desc: m[2].replace(/\s+/g, ' ').trim(),
        qtd: U.parseNumberBR(m[1]) || 1,
        venda: U.parseNumberBR(m[4]),
        ciliaValorLiquido: U.parseNumberBR(m[5])
      });
    }
    return U.normalizeCiliaPieces(pieces);
  };

  U.openApprovalModal = function(os, options) {
    options = options || {};
    const cliente = U.getCliente(os, options.clientes, options.cliente);
    const items = U.buildBudgetItems(os, cliente);
    return new Promise(resolve => {
      if (!items.length) {
        if (typeof options.toast === 'function') options.toast('Nenhum item de orcamento para aprovar.', 'warn');
        resolve(null);
        return;
      }

      let overlay = document.getElementById('modalAprovacaoItensOS');
      if (!overlay) {
        overlay = document.createElement('div');
        overlay.id = 'modalAprovacaoItensOS';
        overlay.className = 'overlay';
        document.body.appendChild(overlay);
      }
      overlay.style.cssText = 'position:fixed;inset:0;z-index:10000;background:rgba(0,0,0,.78);display:none;align-items:center;justify-content:center;padding:16px;backdrop-filter:blur(4px);overflow:auto;';

      const allChecked = options.defaultAll !== false;
      const renderRow = item => `
        <label style="display:grid;grid-template-columns:26px 90px 1fr 110px;gap:10px;align-items:center;padding:10px;border:1px solid rgba(255,255,255,.12);background:rgba(0,0,0,.14);border-radius:4px;margin-bottom:6px;cursor:pointer;">
          <input type="checkbox" class="aprov-item" value="${U.escapeHtml(item.key)}" ${allChecked ? 'checked' : ''} style="width:18px;height:18px;">
          <span style="font-family:var(--fm,var(--mono,monospace));font-size:.68rem;color:${item.tipo === 'peca' ? 'var(--success,#00ff88)' : 'var(--cyan,#00d4ff)'};font-weight:700;">${item.labelTipo}</span>
          <span style="font-size:.82rem;color:var(--text,#e8f4ff);line-height:1.35;">
            ${item.codigo ? `<code style="font-size:.72rem;color:var(--warn,#ffb800);">${U.escapeHtml(item.codigo)}</code> ` : ''}
            ${U.escapeHtml(item.desc || '-')}
            <small style="display:block;color:var(--muted,#7a9ab8);font-family:var(--fm,var(--mono,monospace));font-size:.68rem;margin-top:2px;">
              ${item.tipo === 'servico' ? `Horas/TMO: ${String(item.tempo || 0).replace('.', ',')}h` : `Qtd: ${item.qtd} x ${U.moeda(item.valorUnit)}`}
            </small>
          </span>
          <span style="text-align:right;font-family:var(--fm,var(--mono,monospace));font-weight:700;color:var(--success,#00ff88);">${U.moeda(item.valorFinal)}</span>
        </label>`;

      overlay.innerHTML = `
        <div class="modal" style="max-width:820px;width:96%;max-height:92vh;display:flex;flex-direction:column;background:var(--surf,var(--bg1,#0c1426));border:1px solid var(--border2,var(--border,#24435e));border-radius:6px;color:var(--text,#e8f4ff);box-shadow:0 20px 80px rgba(0,0,0,.45);">
          <div class="modal-head" style="display:flex;align-items:center;justify-content:space-between;gap:12px;padding:14px 18px;border-bottom:1px solid var(--border2,var(--border,#24435e));">
            <div class="modal-title">APROVACAO DO ORCAMENTO - SELECIONE OS ITENS</div>
            <button class="modal-close" type="button" data-aprov-cancel style="width:32px;height:32px;background:transparent;border:1px solid var(--border2,var(--border,#24435e));color:var(--text,#e8f4ff);border-radius:4px;cursor:pointer;">×</button>
          </div>
          <div class="modal-body" style="overflow:auto;padding:18px;">
            <div style="display:flex;gap:8px;flex-wrap:wrap;margin-bottom:12px;">
              <button type="button" class="btn-ghost" data-aprov-all>MARCAR TUDO</button>
              <button type="button" class="btn-ghost" data-aprov-none>DESMARCAR TUDO</button>
            </div>
            <div style="font-size:.78rem;color:var(--muted,#7a9ab8);line-height:1.45;margin-bottom:12px;">
              O orcamento completo sera mantido na O.S. como historico. O financeiro e o fluxo aprovado usarao somente os itens marcados aqui.
            </div>
            ${items.map(renderRow).join('')}
          </div>
          <div class="modal-foot" style="display:flex;justify-content:space-between;gap:10px;align-items:center;padding:12px 18px;border-top:1px solid var(--border2,var(--border,#24435e));flex-wrap:wrap;">
            <div style="font-family:var(--fm,var(--mono,monospace));font-size:.78rem;color:var(--muted,#7a9ab8);" data-aprov-total></div>
            <div style="display:flex;gap:8px;">
              <button class="btn-ghost" type="button" data-aprov-cancel>CANCELAR</button>
              <button class="btn-primary" type="button" data-aprov-confirm>APROVAR SELECIONADOS</button>
            </div>
          </div>
        </div>`;

      function selectedItems() {
        const selected = new Set(Array.from(overlay.querySelectorAll('.aprov-item:checked')).map(i => i.value));
        return items.filter(it => selected.has(it.key));
      }

      function updateTotal() {
        const sel = selectedItems();
        const total = sel.reduce((acc, it) => acc + U.parseNumberBR(it.valorFinal), 0);
        const el = overlay.querySelector('[data-aprov-total]');
        if (el) el.textContent = `${sel.length}/${items.length} item(ns) - Total aprovado: ${U.moeda(total)}`;
      }

      overlay.querySelector('[data-aprov-all]')?.addEventListener('click', () => {
        overlay.querySelectorAll('.aprov-item').forEach(i => { i.checked = true; });
        updateTotal();
      });
      overlay.querySelector('[data-aprov-none]')?.addEventListener('click', () => {
        overlay.querySelectorAll('.aprov-item').forEach(i => { i.checked = false; });
        updateTotal();
      });
      overlay.querySelectorAll('.aprov-item').forEach(i => i.addEventListener('change', updateTotal));
      overlay.querySelectorAll('[data-aprov-cancel]').forEach(btn => btn.addEventListener('click', () => {
        overlay.classList.remove('open');
        overlay.style.display = 'none';
        resolve(null);
      }));
      overlay.querySelector('[data-aprov-confirm]')?.addEventListener('click', () => {
        const sel = selectedItems();
        if (!sel.length) {
          if (typeof options.toast === 'function') options.toast('Selecione ao menos um item aprovado.', 'warn');
          return;
        }
        const total = +sel.reduce((acc, it) => acc + U.parseNumberBR(it.valorFinal), 0).toFixed(2);
        overlay.classList.remove('open');
        overlay.style.display = 'none';
        resolve({
          status: sel.length === items.length ? 'total' : 'parcial',
          totalOrcamento: +items.reduce((acc, it) => acc + U.parseNumberBR(it.valorFinal), 0).toFixed(2),
          totalAprovado: total,
          itens: sel,
          keys: sel.map(it => it.key),
          totalItens: items.length
        });
      });
      updateTotal();
      overlay.classList.add('open');
      overlay.style.display = 'flex';
    });
  };

  U.aprovarOrcamentoComSelecao = async function(options) {
    const db = options?.db || window.db;
    const osId = options?.osId;
    if (!db || !osId) return null;
    const snap = await db.collection('ordens_servico').doc(osId).get();
    if (!snap.exists) throw new Error('O.S. nao encontrada.');
    const os = { id: osId, ...snap.data() };
    const approval = await U.openApprovalModal(os, {
      clientes: options.clientes,
      cliente: options.cliente,
      toast: options.toast || window.toast
    });
    if (!approval) return null;
    const actor = options.actorName || 'Usuario';
    const actorType = options.actorType || 'portal';
    const novoStatus = options.novoStatus || 'Aprovado';
    const timeline = Array.isArray(os.timeline) ? os.timeline.slice() : [];
    timeline.push({
      dt: new Date().toISOString(),
      user: actor,
      acao: `${actor} APROVOU o orcamento (${approval.status}) - ${approval.itens.length}/${approval.totalItens} item(ns) - Total aprovado ${U.moeda(approval.totalAprovado)}`
    });
    const payload = {
      status: novoStatus,
      aprovacao: {
        status: approval.status,
        aprovadoEm: new Date().toISOString(),
        aprovadoPor: actor,
        aprovadoPorTipo: actorType,
        totalOrcamento: approval.totalOrcamento,
        totalAprovado: approval.totalAprovado,
        itens: approval.itens
      },
      itensAprovados: approval.keys,
      totalAprovado: approval.totalAprovado,
      timeline,
      updatedAt: new Date().toISOString()
    };
    await db.collection('ordens_servico').doc(osId).update(payload);
    return payload;
  };

  U.autoDescribeFields = function(root) {
    root = root || document;
    root.querySelectorAll('input, select, textarea, button').forEach(el => {
      if (el.type === 'hidden') return;
      const explicit = el.getAttribute('aria-label') || el.getAttribute('title');
      if (explicit) return;
      let text = '';
      const id = el.id;
      if (id) {
        const label = root.querySelector(`label[for="${CSS.escape(id)}"]`);
        if (label) text = label.textContent.trim();
      }
      if (!text) text = el.closest('.form-group')?.querySelector('label')?.textContent?.trim() || '';
      if (!text) text = el.getAttribute('placeholder') || el.textContent?.trim() || el.name || el.id || '';
      if (text) {
        el.setAttribute('title', text);
        el.setAttribute('aria-label', text);
      }
    });
  };

  window.JarvisOSUtils = U;
  window.JOS = U;

  if (document.readyState === 'loading') {
    document.addEventListener('DOMContentLoaded', () => U.autoDescribeFields(document));
  } else {
    U.autoDescribeFields(document);
  }
})();
