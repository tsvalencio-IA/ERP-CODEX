(function() {
  'use strict';

  const TEMPLATE_URL = 'assets/templates/I-30003_PLANILHA_DE_CUSTOS.xlsx';
  const SERV_START = 19;
  const SERV_END = 74;
  const SERV_TOTAL = 75;
  const PECA_START = 77;
  const PECA_END = 122;
  const PECA_TOTAL = 123;
  const SERV_CAPACITY = SERV_END - SERV_START + 1;
  const PECA_CAPACITY = PECA_END - PECA_START + 1;

  const U = () => window.JarvisOSUtils || window.JOS || {};
  const n = value => U().parseNumberBR ? U().parseNumberBR(value) : (parseFloat(String(value || 0).replace(',', '.')) || 0);

  function moedaNumber(value) {
    return +n(value).toFixed(2);
  }

  function taxaDesconto(value) {
    const pct = n(value);
    if (!pct) return 0;
    return pct > 1 ? +(pct / 100).toFixed(6) : pct;
  }

  function dataExtenso(cidade) {
    const hoje = new Date();
    const meses = ['JANEIRO','FEVEREIRO','MARCO','ABRIL','MAIO','JUNHO','JULHO','AGOSTO','SETEMBRO','OUTUBRO','NOVEMBRO','DEZEMBRO'];
    return `${(cidade || 'SAO PAULO').toUpperCase()}, ${hoje.getDate()} DE ${meses[hoje.getMonth()]} DE ${hoje.getFullYear()}.`;
  }

  function oesNumero(cli, os) {
    const modelo = cli.govOesModelo || 'ORC ###/2026';
    return modelo.replace(/###/g, String(os.id || '').slice(-3).toUpperCase());
  }

  function limparTexto(value) {
    return String(value == null ? '' : value).replace(/\s+/g, ' ').trim();
  }

  function cloneStyle(style) {
    return style ? JSON.parse(JSON.stringify(style)) : {};
  }

  function copiarLinhaModelo(ws, origem, destino, ateColuna) {
    const src = ws.getRow(origem);
    const dst = ws.getRow(destino);
    dst.height = src.height;
    for (let c = 1; c <= ateColuna; c++) {
      const sc = src.getCell(c);
      const dc = dst.getCell(c);
      dc.style = cloneStyle(sc.style);
      dc.numFmt = sc.numFmt;
      dc.alignment = cloneStyle(sc.alignment);
      dc.border = cloneStyle(sc.border);
      dc.fill = cloneStyle(sc.fill);
      dc.font = cloneStyle(sc.font);
      dc.protection = cloneStyle(sc.protection);
    }
  }

  function setCell(ws, addr, value) {
    const c = ws.getCell(addr);
    c.value = value == null ? '' : value;
  }

  function setFormula(ws, addr, formula, result) {
    const c = ws.getCell(addr);
    c.value = { formula, result: moedaNumber(result) };
  }

  function setNumberCell(ws, addr, value, numFmt) {
    const c = ws.getCell(addr);
    c.value = moedaNumber(value);
    if (numFmt) c.numFmt = numFmt;
    c.alignment = { ...(c.alignment || {}), horizontal: 'right', vertical: 'middle', shrinkToFit: true };
  }

  function setMoneyCell(ws, addr, value) {
    setNumberCell(ws, addr, value, '"R$" #,##0.00;-"R$" #,##0.00;"-"');
  }

  function setPercentCell(ws, addr, value) {
    const c = ws.getCell(addr);
    c.value = n(value) || 0;
    c.numFmt = '0.0%';
    c.alignment = { ...(c.alignment || {}), horizontal: 'center', vertical: 'middle', shrinkToFit: true };
  }

  function safeUnmerge(ws, range) {
    try { ws.unMergeCells(range); } catch(e) {}
  }

  function limparRangeResumo(ws, row) {
    ['A:H','A:D','A:B','B:D','B:F','B:H','C:D','D:H','E:F','E:H','F:H','G:H'].forEach(range => safeUnmerge(ws, `${range.split(':')[0]}${row}:${range.split(':')[1]}${row}`));
    ['A','B','C','D','E','F','G','H'].forEach(col => setCell(ws, col + row, ''));
  }

  function inserirLinhasExtras(ws, totalRow, modeloRow, extra) {
    if (extra <= 0) return;
    ws.spliceRows(totalRow, 0, ...Array.from({ length: extra }, () => []));
    for (let i = 0; i < extra; i++) copiarLinhaModelo(ws, modeloRow, totalRow + i, 8);
  }

  function enderecoOficina(tenant) {
    return tenant.enderecoCompleto || [tenant.endereco, tenant.numero, tenant.bairro, tenant.cidade, tenant.uf, tenant.cep].filter(Boolean).join(', ');
  }

  function aplicarAjustesCabecalho(ws) {
    ws.getColumn('B').width = 26;
    ws.getColumn('D').width = 46;
    ws.getColumn('E').width = 8;
    ws.getColumn('F').width = 17;
    ws.getColumn('G').width = 12;
    ws.getColumn('H').width = 20;
    [1,3,5,6,7,9,10,11,12,14,15,17].forEach(r => {
      const row = ws.getRow(r);
      row.height = Math.max(row.height || 15, r === 1 ? 92 : 18);
      row.eachCell({ includeEmpty: true }, cell => {
        cell.alignment = { ...(cell.alignment || {}), wrapText: true, shrinkToFit: true, vertical: 'middle' };
      });
    });
  }

  function clearRow(ws, row, cols) {
    (cols || ['B','D','E','F','G','H']).forEach(col => setCell(ws, col + row, ''));
  }

  function prepararLinhasDados(ws, start, end, usadas) {
    const visiveis = Math.min(end - start + 1, Math.max(usadas + 3, 4));
    const ultimoVisivel = start + visiveis - 1;
    for (let r = start; r <= end; r++) {
      clearRow(ws, r);
      ws.getRow(r).hidden = r > ultimoVisivel;
    }
  }

  function prepararRodape(ws, rows) {
    rows.forEach(r => {
      ws.getRow(r).hidden = false;
      limparRangeResumo(ws, r);
    });
  }

  function estilizarResumo(ws, row, destaque) {
    const fill = destaque ? { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFC9C19B' } } : { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFEDEDED' } };
    ['B','C','D','E','F','G','H'].forEach(col => {
      const cell = ws.getCell(col + row);
      cell.fill = fill;
      cell.border = {
        top: { style: 'thin', color: { argb: 'FF000000' } },
        left: { style: 'thin', color: { argb: 'FF000000' } },
        bottom: { style: 'thin', color: { argb: 'FF000000' } },
        right: { style: 'thin', color: { argb: 'FF000000' } }
      };
      cell.alignment = { ...(cell.alignment || {}), vertical: 'middle', shrinkToFit: true, wrapText: true };
      cell.font = { ...(cell.font || {}), bold: true, color: { argb: 'FF000000' }, size: destaque ? 15 : 10 };
    });
    ws.getRow(row).height = destaque ? 26 : 19;
  }

  function linhaResumoValor(ws, row, label, value, opts) {
    const op = opts || {};
    limparRangeResumo(ws, row);
    safeUnmerge(ws, `B${row}:F${row}`);
    ws.mergeCells(`B${row}:F${row}`);
    setCell(ws, 'B' + row, label);
    ws.getCell('B' + row).alignment = { horizontal: op.contrato ? 'center' : 'left', vertical: 'middle', wrapText: true, shrinkToFit: true };
    if (op.blankValue) {
      setCell(ws, 'G' + row, '');
      setCell(ws, 'H' + row, '');
    } else {
      setCell(ws, 'G' + row, 'R$');
      ws.getCell('G' + row).alignment = { horizontal: 'center', vertical: 'middle' };
      setMoneyCell(ws, 'H' + row, value);
    }
    estilizarResumo(ws, row, !!op.contrato);
  }

  function linhaTotalServicos(ws, row, horas, total) {
    limparRangeResumo(ws, row);
    setCell(ws, 'B' + row, 'TOTAL DE SERVICOS');
    setNumberCell(ws, 'E' + row, horas, '0.00');
    setCell(ws, 'G' + row, 'R$');
    ws.getCell('G' + row).alignment = { horizontal: 'center', vertical: 'middle' };
    setMoneyCell(ws, 'H' + row, total);
    estilizarResumo(ws, row, false);
  }

  function linhaTotalPecas(ws, row, total) {
    limparRangeResumo(ws, row);
    safeUnmerge(ws, `B${row}:F${row}`);
    ws.mergeCells(`B${row}:F${row}`);
    setCell(ws, 'B' + row, 'TOTAL DE PECAS');
    setCell(ws, 'G' + row, 'R$');
    ws.getCell('G' + row).alignment = { horizontal: 'center', vertical: 'middle' };
    setMoneyCell(ws, 'H' + row, total);
    estilizarResumo(ws, row, false);
  }

  function formatarCabecalhoPM(texto) {
    const linhas = String(texto || '')
      .split(/\r?\n/)
      .map(l => limparTexto(l))
      .filter(Boolean);
    return linhas.map(l => '                                      ' + l).join('\n');
  }

  function coletarDados(os, cli, veiculo) {
    const sessao = window.J || {};
    const tenant = { ...sessao, ...(sessao.oficina || {}) };
    tenant.tnome = tenant.nomeFantasia || sessao.tnome || tenant.tnome || '';
    tenant.nome = tenant.nome || sessao.nome || tenant.tnome || '';
    const servicos = (os.servicos || []).filter(s => s.desc || s.valor || s.tempo);
    const pecas = (os.pecas || []).filter(p => p.desc || p.descricao || p.codigo || p.cod || p.venda || p.valor);
    const descMO = taxaDesconto(os.descMO != null ? os.descMO : cli.govDescMO);
    const descPeca = taxaDesconto(os.descPeca != null ? os.descPeca : cli.govDescPeca);
    const valorHoraCliente = n(cli.govValorHora || 0);

    const linhasServ = servicos.map(s => {
      const tempo = n(s.tempo || 0);
      const valorBrutoServico = n(s.valor || 0);
      const resolvido = U().resolvePMSPServico ? U().resolvePMSPServico(s, { veiculo, fallbackValorHora: valorHoraCliente }) : {};
      const valorHora = n(s.valorHora || s.valorHoraSecao || resolvido.valorHora || 0) ||
        (tempo > 0 && valorBrutoServico > 0 ? +(valorBrutoServico / tempo).toFixed(2) : 0) ||
        valorHoraCliente;
      const sistemaServico = resolvido.secaoHoraLabel || s.secaoHoraLabel || s.sistemaTabela || s.sistema || '';
      const totalFinal = +(valorHora * tempo * (1 - descMO)).toFixed(2);
      return {
        sistema: limparTexto(sistemaServico),
        desc: limparTexto(s.desc || ''),
        tempo,
        valorHora,
        descPct: descMO,
        total: totalFinal
      };
    });

    const linhasPecas = pecas.map(p => {
      const qtd = n(p.qtd || p.q || 1) || 1;
      const valorUnit = n(p.venda || p.valor || p.v);
      const totalFinal = +(qtd * valorUnit * (1 - descPeca)).toFixed(2);
      return {
        codigo: limparTexto(p.codigo || p.cod || 'sem oem') || 'sem oem',
        desc: limparTexto(p.desc || p.descricao || ''),
        qtd,
        valorUnit,
        descPct: descPeca,
        total: totalFinal
      };
    });

    return { tenant, linhasServ, linhasPecas, descMO, descPeca };
  }

  async function exportarComExcelJS(os, cli, veiculo) {
    if (typeof ExcelJS === 'undefined') return false;

    const resp = await fetch(TEMPLATE_URL, { cache: 'no-store' });
    if (!resp.ok) throw new Error('Modelo PMSP nao encontrado: ' + TEMPLATE_URL);
    const buffer = await resp.arrayBuffer();
    const wb = new ExcelJS.Workbook();
    await wb.xlsx.load(buffer);
    const ws = wb.worksheets[0];

    const { tenant, linhasServ, linhasPecas } = coletarDados(os, cli, veiculo);

    const servRowsWanted = Math.max(linhasServ.length + 3, 4);
    const servExtra = Math.max(0, servRowsWanted - SERV_CAPACITY);
    inserirLinhasExtras(ws, SERV_TOTAL, SERV_END, servExtra);
    const servEnd = SERV_END + servExtra;
    const servTotal = SERV_TOTAL + servExtra;
    const pecaShift = servExtra;
    const pecaStart = PECA_START + pecaShift;
    const pecaEndBase = PECA_END + pecaShift;
    const pecaTotalBase = PECA_TOTAL + pecaShift;

    const pecaRowsWanted = Math.max(linhasPecas.length + 3, 4);
    const pecaExtra = Math.max(0, pecaRowsWanted - PECA_CAPACITY);
    inserirLinhasExtras(ws, pecaTotalBase, pecaEndBase, pecaExtra);
    const pecaEnd = pecaEndBase + pecaExtra;
    const pecaTotal = pecaTotalBase + pecaExtra;
    const totalGeralRow = pecaTotal + 1;
    const vistoriaRow = pecaTotal + 2;
    const resumoPecasRow = pecaTotal + 3;
    const resumoMORow = pecaTotal + 4;
    const resumoTotalRow = pecaTotal + 5;
    const contratoRow = pecaTotal + 6;
    const dataRow = pecaTotal + 7;
    const representanteRow = pecaTotal + 11;

    const cabecalho = (cli.govCabecalho || '').trim() || 'SECRETARIA DA SEGURANCA PUBLICA\nPOLICIA MILITAR DO ESTADO DE SAO PAULO';
    const razaoOficina = tenant.razaoSocial || tenant.nomeFantasia || tenant.tnome || '';
    const telefoneOficina = tenant.telefone || tenant.wpp || '';
    const representante = tenant.representante || tenant.responsavel || tenant.orcamentista || tenant.nome || '';
    setCell(ws, 'B1', formatarCabecalhoPM(cabecalho));
    setCell(ws, 'A3', `REFERENCIA: ORDEM E EXECUCAO DE SERVICOS No ${oesNumero(cli, os)}`);
    setCell(ws, 'A5', `MARCA: ${(veiculo.marca || '').toUpperCase()}`);
    setCell(ws, 'C5', `MODELO: ${(veiculo.modelo || '').toUpperCase()}`);
    setCell(ws, 'E5', `ANO: ${veiculo.ano || ''}`);
    setCell(ws, 'G5', `PLACA: ${(veiculo.placa || os.placa || '').toUpperCase()}`);
    setCell(ws, 'A6', `CHASSIS: ${(veiculo.chassis || '').toUpperCase()}`);
    setCell(ws, 'D6', `PATRIMONIO: ${veiculo.patrimonio || ''}`);
    setCell(ws, 'A7', `KM: ${os.km || veiculo.km || ''}`);
    setCell(ws, 'C7', `PREFIXO: ${(veiculo.prefixo || '').toUpperCase()}`);
    setCell(ws, 'E7', `OPM DETENTORA: ${cli.govUnidade || cli.nome || ''}`);
    setCell(ws, 'A9', `RAZAO SOCIAL : ${razaoOficina}`);
    setCell(ws, 'E9', `CNPJ: ${tenant.cnpj || ''}`);
    setCell(ws, 'A10', `ENDERECO: ${enderecoOficina(tenant)}`);
    setCell(ws, 'A11', `TELEFONE: ${telefoneOficina}`);
    setCell(ws, 'D11', `ORCAMENTISTA: ${tenant.orcamentista || tenant.nome || ''}`);
    setCell(ws, 'A12', `REPRESENTANTE LEGAL: ${representante}`);
    setCell(ws, 'A14', `UNIDADE : ${cli.govUnidade || cli.nome || ''}`);
    setCell(ws, 'E14', `CNPJ: ${cli.doc || ''}`);
    setCell(ws, 'A15', `ENDERECO: ${[cli.rua, cli.num, cli.bairro, cli.cidade].filter(Boolean).join(', ')}`);
    setCell(ws, 'A17', `FISCAL DO CONTRATO: ${cli.govFiscal || ''}`);
    aplicarAjustesCabecalho(ws);

    prepararLinhasDados(ws, SERV_START, servEnd, linhasServ.length);
    linhasServ.forEach((s, idx) => {
      const r = SERV_START + idx;
      ws.getRow(r).hidden = false;
      setCell(ws, 'B' + r, s.sistema);
      setCell(ws, 'D' + r, s.desc);
      setCell(ws, 'E' + r, s.tempo);
      setMoneyCell(ws, 'F' + r, s.valorHora);
      setPercentCell(ws, 'G' + r, s.descPct);
      setMoneyCell(ws, 'H' + r, s.total);
    });

    prepararLinhasDados(ws, pecaStart, pecaEnd, linhasPecas.length);
    linhasPecas.forEach((p, idx) => {
      const r = pecaStart + idx;
      ws.getRow(r).hidden = false;
      setCell(ws, 'B' + r, p.codigo);
      setCell(ws, 'D' + r, p.desc);
      setCell(ws, 'E' + r, p.qtd);
      setMoneyCell(ws, 'F' + r, p.valorUnit);
      setPercentCell(ws, 'G' + r, p.descPct);
      setMoneyCell(ws, 'H' + r, p.total);
    });

    const totalPecas = linhasPecas.reduce((sum, p) => sum + p.total, 0);
    const totalMO = linhasServ.reduce((sum, s) => sum + s.total, 0);
    const totalHoras = linhasServ.reduce((sum, s) => sum + s.tempo, 0);
    const contrato = +(totalPecas + totalMO).toFixed(2);
    prepararRodape(ws, [
      servTotal,
      pecaTotal,
      totalGeralRow,
      vistoriaRow,
      resumoPecasRow,
      resumoMORow,
      resumoTotalRow,
      contratoRow,
      dataRow,
      representanteRow
    ]);
    linhaTotalServicos(ws, servTotal, totalHoras, totalMO);
    linhaTotalPecas(ws, pecaTotal, totalPecas);
    linhaResumoValor(ws, totalGeralRow, 'TOTAL GERAL', 0, { blankValue: true });
    linhaResumoValor(ws, vistoriaRow, 'VALOR DA VISTORIA TECNICA COMPLEMENTAR AO ESCOPO DE SERVICOS', 0, { blankValue: true });
    linhaResumoValor(ws, resumoPecasRow, 'VALOR TOTAL DE PECAS', totalPecas);
    linhaResumoValor(ws, resumoMORow, 'VALOR TOTAL DE MAO DE OBRA', totalMO);
    linhaResumoValor(ws, resumoTotalRow, 'TOTAL GERAL', contrato);
    linhaResumoValor(ws, contratoRow, 'VALOR DO CONTRATO', contrato, { contrato: true });
    setCell(ws, 'A' + dataRow, dataExtenso(tenant.cidade || cli.cidade));
    setCell(ws, 'A' + representanteRow, String(representante).toUpperCase());

    ws.pageSetup = {
      ...ws.pageSetup,
      orientation: 'portrait',
      fitToPage: true,
      fitToWidth: 1,
      fitToHeight: 0,
      printArea: `A1:H${representanteRow}`
    };

    const fname = `${(veiculo.prefixo || os.id.slice(-6).toUpperCase())}_PLANILHA_DE_CUSTOS.xlsx`;
    const out = await wb.xlsx.writeBuffer();
    const blob = new Blob([out], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
    const a = document.createElement('a');
    a.href = URL.createObjectURL(blob);
    a.download = fname;
    document.body.appendChild(a);
    a.click();
    a.remove();
    setTimeout(() => URL.revokeObjectURL(a.href), 1000);
    window.toast?.(`Orcamento PMSP exportado: ${fname}`, 'ok');
    return true;
  }

  async function exportarFallbackSheetJS(os, cli, veiculo) {
    if (typeof XLSX === 'undefined') throw new Error('Biblioteca XLSX nao carregou.');
    const { tenant, linhasServ, linhasPecas } = coletarDados(os, cli, veiculo);
    const rows = [
      ['SECRETARIA DA SEGURANCA PUBLICA POLICIA MILITAR DO ESTADO DE SAO PAULO'],
      ['PLANILHA DE COMPOSICAO DE CUSTOS'],
      [`REFERENCIA: ORDEM E EXECUCAO DE SERVICOS No ${oesNumero(cli, os)}`],
      ['DADOS DA VIATURA'],
      [`MARCA: ${veiculo.marca || ''}`, `MODELO: ${veiculo.modelo || ''}`, `ANO: ${veiculo.ano || ''}`, `PLACA: ${veiculo.placa || os.placa || ''}`],
      [`CHASSIS: ${veiculo.chassis || ''}`, `PATRIMONIO: ${veiculo.patrimonio || ''}`],
      [`KM: ${os.km || veiculo.km || ''}`, `PREFIXO: ${veiculo.prefixo || ''}`, `OPM DETENTORA: ${cli.govUnidade || cli.nome || ''}`],
      ['DADOS DA EMPRESA'],
      [`RAZAO SOCIAL: ${tenant.razaoSocial || tenant.nomeFantasia || tenant.tnome || ''}`, `CNPJ: ${tenant.cnpj || ''}`],
      [`ENDERECO: ${enderecoOficina(tenant)}`],
      [`TELEFONE: ${tenant.telefone || tenant.wpp || ''}`, `ORCAMENTISTA: ${tenant.orcamentista || tenant.nome || ''}`],
      [`REPRESENTANTE LEGAL: ${tenant.representante || tenant.responsavel || tenant.nome || ''}`],
      ['DADOS DO CLIENTE'],
      [`UNIDADE: ${cli.govUnidade || cli.nome || ''}`, `CNPJ: ${cli.doc || ''}`],
      [`ENDERECO: ${[cli.rua, cli.num, cli.bairro, cli.cidade].filter(Boolean).join(', ')}`],
      [`FISCAL DO CONTRATO: ${cli.govFiscal || ''}`],
      [],
      ['DESCRICAO DO SISTEMA','DESCRICAO DO SERVICO','TMO','VALOR','DESC.','VALOR']
    ];
    linhasServ.forEach(s => rows.push([s.sistema, s.desc, s.tempo, s.valorHora, s.descPct, s.total]));
    rows.push(['TOTAL DE SERVICOS', '', linhasServ.reduce((sum, s) => sum + s.tempo, 0), '', '', linhasServ.reduce((sum, s) => sum + s.total, 0)]);
    rows.push([]);
    rows.push(['CODIGO DA PECA (CODIGO ORIGINAL)','DESCRICAO','QTD','VALOR UNITARIO REGISTRADO','DESC','VALOR']);
    linhasPecas.forEach(p => rows.push([p.codigo, p.desc, p.qtd, p.valorUnit, p.descPct, p.total]));
    rows.push(['TOTAL DE PECAS', '', '', '', '', linhasPecas.reduce((sum, p) => sum + p.total, 0)]);
    const total = linhasPecas.reduce((sum, p) => sum + p.total, 0) + linhasServ.reduce((sum, s) => sum + s.total, 0);
    rows.push(['VALOR DO CONTRATO', '', '', '', '', total]);
    rows.push([dataExtenso(tenant.cidade || cli.cidade)]);

    const ws = XLSX.utils.aoa_to_sheet(rows);
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, 'Plan1');
    const fname = `${(veiculo.prefixo || os.id.slice(-6).toUpperCase())}_PLANILHA_DE_CUSTOS.xlsx`;
    XLSX.writeFile(wb, fname);
    window.toast?.('ExcelJS nao carregou; exportado em modo compatibilidade.', 'warn');
  }

  window.exportarOrcamentoPMSP = async function() {
    try {
      const osId = document.getElementById('osId')?.value;
      if (!osId) { window.toast?.('Salve a O.S. antes de exportar.', 'warn'); return; }

      const os = (window.J?.os || []).find(o => o.id === osId);
      if (!os) { window.toast?.('O.S. nao encontrada.', 'err'); return; }

      const cli = (window.J?.clientes || []).find(c => c.id === os.clienteId);
      if (!cli || cli.tipoCliente !== 'governo') {
        window.toast?.('Esta exportacao e exclusiva para clientes governamentais.', 'err');
        return;
      }

      const veiculo = (window.J?.veiculos || []).find(v => v.id === os.veiculoId) || {};
      const ok = await exportarComExcelJS(os, cli, veiculo);
      if (!ok) await exportarFallbackSheetJS(os, cli, veiculo);
    } catch (e) {
      console.error('[PMSP XLSX]', e);
      window.toast?.('Erro ao exportar PMSP: ' + e.message, 'err');
    }
  };
})();
