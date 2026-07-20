// ==UserScript==
// @name         Exportar Conclusões Projudi para Excel
// @namespace    https://projudi2.tjpr.jus.br/
// @version      7.1
// @description  Coleta conclusões (remessa), retorno de conclusos ou juntadas, acumula e exporta em Excel ou PDF (com painel e gráficos)
// @author       rcpleme2
// @match        https://projudi2.tjpr.jus.br/projudi/processo/conclusao.do*
// @match        https://projudi2.tjpr.jus.br/projudi/processo/analisarJuntada.do*
// @require      https://cdn.sheetjs.com/xlsx-latest/package/dist/xlsx.full.min.js
// @require      https://cdnjs.cloudflare.com/ajax/libs/jspdf/2.5.1/jspdf.umd.min.js
// @require      https://cdnjs.cloudflare.com/ajax/libs/jspdf-autotable/3.8.2/jspdf.plugin.autotable.min.js
// @grant        GM_addStyle
// @grant        GM_download
// ==/UserScript==

(function () {
    'use strict';

    const store = window.localStorage;
    const STALE_MS = 2 * 60 * 1000; // 2 minutos sem atividade => coleta em andamento considerada obsoleta

    // ── Leitura da Atuação atual ────────────────────────────────────────────────

    function lerAtuacao() {
        // Localiza o bloco cujo rótulo é "Atuação:" (não o primeiro div.group, que é "Atribuição:")
        const grupos = document.querySelectorAll('div.group');
        for (const grupo of grupos) {
            const label = grupo.querySelector('span.userinfo_label');
            if (label && /atua[çc][ãa]o/i.test(label.textContent)) {
                const span = grupo.querySelector('span[title]');
                return span ? span.textContent.trim() : '';
            }
        }
        return '';
    }

    function textoCelula(td) {
        return td ? (td.textContent || '').trim() : '';
    }

    function norm(s) {
        return (s || '').replace(/\s+/g, ' ').trim();
    }

    // Texto do elemento ignorando o conteúdo dos filhos <b> (ex.: separar a
    // especificação — em texto normal — do tipo de documento — em negrito).
    function textoSemNegrito(el) {
        if (!el) return '';
        let s = '';
        el.childNodes.forEach(n => {
            if (n.nodeType === 1 && n.tagName === 'B') return;
            s += n.textContent;
        });
        return norm(s);
    }

    function processoEclasse(td) {
        const em = td.querySelector('em');
        const div = td.querySelector('div');
        return {
            processo: em ? em.textContent.trim() : textoCelula(td),
            classe: div ? div.textContent.trim() : '',
        };
    }

    // ── Configurações dos dois relatórios ───────────────────────────────────────
    // Ambos ficam em conclusao.do; distinguem-se pelo cabeçalho da tabela:
    //   "Dt. Remessa"  => relatório de Conclusões (remessa)      — 9 colunas
    //   "Dt. Retorno"  => relatório de Retorno de Conclusos       — 10 colunas (1ª = semáforo)

    const CFG_CONCLUSOES = {
        prefixo: 'projudi_export_',            // mantém as chaves já usadas (dados existentes preservados)
        detecta: (cab) => /remessa/i.test(cab),
        minTds: 8,
        usaAtuacao: true,
        nomeArquivo: 'conclusoes_projudi',
        rotulos: { coletar: 'Coletar esta Atuação', coletarMais: 'Coletar mais uma Atuação', baixar: '⬇ Baixar planilha' },
        cabecalhos: ['Atuação', 'Dt. Remessa', 'Processo', 'Classe', 'Seq.', 'Tipo de Conclusão', 'Privativa', 'Responsável', 'Pré-análise', 'Agrupador'],
        larguras: [{ wch: 24 }, { wch: 18 }, { wch: 26 }, { wch: 18 }, { wch: 6 }, { wch: 30 }, { wch: 10 }, { wch: 30 }, { wch: 14 }, { wch: 20 }],
        extrai: (tds, atuacao) => {
            const pc = processoEclasse(tds[2]);
            return {
                atuacao,
                dtRemessa: textoCelula(tds[1]),
                processo: pc.processo,
                classe: pc.classe,
                seq: textoCelula(tds[3]),
                tipoConclusao: textoCelula(tds[4]),
                privativa: textoCelula(tds[5]),
                responsavel: textoCelula(tds[6]),
                preAnalise: textoCelula(tds[7]),
                agrupador: textoCelula(tds[8]),
            };
        },
        linha: (d) => [d.atuacao, d.dtRemessa, d.processo, d.classe, d.seq, d.tipoConclusao, d.privativa, d.responsavel, d.preAnalise, d.agrupador],
    };

    const CFG_RETORNO = {
        prefixo: 'projudi_retorno_',
        detecta: (cab) => /retorno/i.test(cab),
        minTds: 9,                              // linhas têm 10 tds (0 = semáforo)
        usaAtuacao: false,
        nomeArquivo: 'retorno_conclusos_projudi',
        rotulos: { coletar: 'Extrair Retorno', coletarMais: 'Extrair mais (Retorno)', baixar: '⬇ Baixar Retorno' },
        // Colunas pedidas: Processo - Classe - Dt. Retorno - Tipo de conclusão - Responsável - Agrupador
        cabecalhos: ['Processo', 'Classe', 'Dt. Retorno', 'Tipo de Conclusão', 'Responsável', 'Agrupador'],
        larguras: [{ wch: 26 }, { wch: 18 }, { wch: 18 }, { wch: 30 }, { wch: 30 }, { wch: 20 }],
        extrai: (tds) => {
            const pc = processoEclasse(tds[3]);
            return {
                processo: pc.processo,
                classe: pc.classe,
                dtRetorno: textoCelula(tds[2]),
                tipoConclusao: textoCelula(tds[5]),
                responsavel: textoCelula(tds[7]),
                agrupador: textoCelula(tds[9]),
            };
        },
        linha: (d) => [d.processo, d.classe, d.dtRetorno, d.tipoConclusao, d.responsavel, d.agrupador],
        pdf: {
            titulo: 'Retorno de Processos Conclusos',
            atosTitulo: 'Retornos de conclusão pendentes',
            agingTitulo: 'Retornos por tempo de espera',
            dataCampo: 'dtRetorno',
            dataTitulo: 'Retorno pendente mais antigo',
            processoCampo: 'processo',
            tipoCampo: 'tipoConclusao',
            grupoCampo: 'agrupador',
            grupoTitulo: 'Processos por Agrupador',
            colunas: [
                { header: 'Processo', width: 34, get: (d) => d.processo },
                { header: 'Classe', width: 24, get: (d) => d.classe },
                { header: 'Dt. Retorno', width: 26, get: (d) => d.dtRetorno },
                { header: 'Tipo de Conclusão', width: 46, get: (d) => d.tipoConclusao },
                { header: 'Responsável', width: 46, get: (d) => d.responsavel },
                { header: 'Agrupador', width: 40, get: (d) => d.agrupador },
                { header: 'Dias', width: 14, get: (d, ctx) => diasDecorridos(d.dtRetorno, ctx.now) },
            ],
        },
    };

    const CFG_JUNTADAS = {
        prefixo: 'projudi_juntadas_',
        detecta: (cab) => /juntado\s+por/i.test(cab),
        minTds: 9,                              // linhas têm 10 tds (0=checkbox, 1=expandir, 2=semáforo)
        usaAtuacao: false,
        nomeArquivo: 'juntadas_projudi',
        rotulos: { coletar: 'Extrair Juntadas', coletarMais: 'Extrair mais (Juntadas)', baixar: '⬇ Baixar Juntadas' },
        // Colunas pedidas: Processo - Tipo de documento - Especificação - Data de Envio - Juntado por
        cabecalhos: ['Processo', 'Tipo de Documento', 'Especificação do Documento', 'Data de Envio', 'Juntado por'],
        larguras: [{ wch: 26 }, { wch: 40 }, { wch: 52 }, { wch: 18 }, { wch: 28 }],
        extrai: (tds) => {
            // Processo (td[3]): número em <em class="attention">
            const emProc = tds[3].querySelector('em');
            const processo = emProc ? emProc.textContent.trim() : textoCelula(tds[3]);

            // Tipo de Documento (td[6]): título em <b>; especificação = texto após o <br>
            const tdDoc = tds[6];
            const b = tdDoc.querySelector('b');
            const emDoc = tdDoc.querySelector('em') || tdDoc;
            const tipoDocumento = b ? norm(b.textContent) : textoCelula(tdDoc);
            const especificacao = textoSemNegrito(emDoc);

            // "Juntado por" (td[8]): nome (texto) + função (em <b>, ex.: "Procurador")
            const bFunc = tds[8].querySelector('b');
            return {
                processo,
                tipoDocumento,
                especificacao,
                dataEnvio: textoCelula(tds[7]),
                juntadoPor: textoSemNegrito(tds[8]), // nome, sem o cargo em negrito
                funcao: bFunc ? norm(bFunc.textContent) : '',
            };
        },
        linha: (d) => [d.processo, d.tipoDocumento, d.especificacao, d.dataEnvio, d.juntadoPor],
        pdf: {
            titulo: 'Juntadas',
            atosTitulo: 'Juntadas pendentes',
            agingTitulo: 'Juntadas por tempo de espera',
            dataCampo: 'dataEnvio',
            dataTitulo: 'Juntada pendente mais antiga',
            processoCampo: 'processo',
            tipoCampo: 'tipoDocumento',
            grupoCampo: 'tipoDocumento',
            grupoTitulo: 'Processos por Tipo de Documento',
            colunas: [
                { header: 'Processo', width: 34, get: (d) => d.processo },
                { header: 'Tipo de Documento', width: 40, get: (d) => d.tipoDocumento },
                { header: 'Especificação do Documento', width: 66, get: (d) => d.especificacao },
                { header: 'Data de Envio', width: 26, get: (d) => d.dataEnvio },
                { header: 'Juntado por', width: 36, get: (d) => d.juntadoPor },
                { header: 'Função', width: 26, get: (d) => d.funcao || '' },
                { header: 'Dias', width: 13, get: (d, ctx) => diasDecorridos(d.dataEnvio, ctx.now) },
            ],
        },
    };

    // ── Coletor genérico (parametrizado por configuração) ───────────────────────

    function criarColetor(cfg) {
        const KEY_RODANDO     = cfg.prefixo + 'rodando';
        const KEY_NUM_PAGINAS = cfg.prefixo + 'num_paginas';
        const KEY_PAGINA_PREF = cfg.prefixo + 'pagina_';
        const KEY_TS          = cfg.prefixo + 'ts';

        function marcarAtividade() { store.setItem(KEY_TS, String(Date.now())); }
        function obsoleta() {
            const ts = parseInt(store.getItem(KEY_TS) || '0', 10);
            return !ts || (Date.now() - ts) > STALE_MS;
        }
        function rodando() { return store.getItem(KEY_RODANDO) === '1'; }

        function limparTudo() {
            const n = parseInt(store.getItem(KEY_NUM_PAGINAS) || '0', 10);
            for (let i = 0; i < n; i++) store.removeItem(KEY_PAGINA_PREF + i);
            store.removeItem(KEY_NUM_PAGINAS);
            store.removeItem(KEY_RODANDO);
            store.removeItem(KEY_TS);
        }

        function adicionarPagina(dadosPagina) {
            const idx = parseInt(store.getItem(KEY_NUM_PAGINAS) || '0', 10);
            store.setItem(KEY_PAGINA_PREF + idx, JSON.stringify(dadosPagina));
            store.setItem(KEY_NUM_PAGINAS, String(idx + 1));
            return contarRegistros();
        }

        function normalizarParaArray(valor) {
            let v = valor, t = 0;
            while (typeof v === 'string' && t < 5) {
                try { v = JSON.parse(v); } catch (e) { break; }
                t++;
            }
            return Array.isArray(v) ? v : null;
        }

        function lerTudo() {
            const n = parseInt(store.getItem(KEY_NUM_PAGINAS) || '0', 10);
            let dados = [];
            for (let i = 0; i < n; i++) {
                const bruto = store.getItem(KEY_PAGINA_PREF + i);
                if (!bruto) continue;
                const parte = normalizarParaArray(bruto);
                if (parte) dados = dados.concat(parte);
                else console.error('[Exportar Projudi] parte ilegível no índice', i);
            }
            return dados;
        }

        function contarRegistros() { return lerTudo().length; }
        function contarAtuacoes() {
            const s = new Set(lerTudo().map(d => d.atuacao || ''));
            s.delete('');
            return s.size;
        }

        function coletarPaginaAtual() {
            const atuacao = cfg.usaAtuacao ? lerAtuacao() : '';
            const linhas = document.querySelectorAll('table.resultTable tbody tr');
            const dados = [];
            linhas.forEach(tr => {
                const tds = tr.querySelectorAll('td');
                if (tds.length < cfg.minTds) return;
                dados.push(cfg.extrai(tds, atuacao));
            });
            return dados;
        }

        function iniciar() {
            store.setItem(KEY_RODANDO, '1');
            marcarAtividade();
            continuar();
        }

        function continuar() {
            desabilitarBotoes(true);
            marcarAtividade();

            const dadosPagina = coletarPaginaAtual();
            let total;
            try {
                total = adicionarPagina(dadosPagina);
            } catch (err) {
                store.removeItem(KEY_RODANDO);
                atualizarStatus(`Erro ao armazenar dados (${err.name}). Os dados já coletados foram mantidos.`);
                console.error('[Exportar Projudi]', err);
                render();
                return;
            }

            const pagina = numeroPaginaAtual();
            const totReg = totalRegistrosPagina();
            const porPag = dadosPagina.length || 20;
            const totPag = totReg ? Math.ceil(totReg / porPag) : '?';
            const ctx = cfg.usaAtuacao ? `"${lerAtuacao() || '(sem atuação)'}" — ` : '';

            atualizarStatus(`Coletando ${ctx}página ${pagina} de ${totPag} — ${total} no total acumulado...`);

            if (temProximaPagina()) {
                document.querySelector('a.arrowNextOn').click();
            } else {
                store.removeItem(KEY_RODANDO);
                store.removeItem(KEY_TS);
                render();
                const extra = cfg.usaAtuacao ? ` de ${contarAtuacoes()} atuação(ões)` : '';
                const dica = cfg.usaAtuacao ? ' Troque de atuação e colete mais, ou baixe a planilha.' : ' Baixe a planilha ou colete mais.';
                atualizarStatus(`Coleta concluída. Acumulado: ${total} registros${extra}.${dica}`);
            }
        }

        function baixar() {
            try {
                const dados = lerTudo();
                if (!dados.length) { atualizarStatus('Nenhum registro coletado para exportar.'); return; }
                gerarEbaixarExcel(dados, cfg);
                const extra = cfg.usaAtuacao ? ` de ${contarAtuacoes()} atuação(ões)` : '';
                atualizarStatus(`✓ ${dados.length} registros${extra} exportados. Use "Limpar" para começar de novo.`);
            } catch (err) {
                atualizarStatus(`Erro ao gerar planilha: ${err.message}`);
                console.error('[Exportar Projudi]', err);
            }
        }

        function pdf() {
            try {
                const dados = lerTudo();
                if (!dados.length) { atualizarStatus('Nenhum registro coletado para exportar.'); return; }
                gerarPDF(dados, cfg);
                atualizarStatus(`✓ PDF gerado com ${dados.length} registros.`);
            } catch (err) {
                atualizarStatus(`Erro ao gerar PDF: ${err.message}`);
                console.error('[Exportar Projudi]', err);
            }
        }

        function limpar() {
            limparTudo();
            render();
            atualizarStatus('Dados acumulados apagados. Pronto para uma nova coleta.');
        }

        function render() {
            const total = contarRegistros();
            const bColetar = document.getElementById('btn-coletar');
            const bBaixar = document.getElementById('btn-baixar');
            const bLimpar = document.getElementById('btn-limpar');
            if (!bColetar) return;

            bColetar.disabled = false;
            bColetar.textContent = total > 0 ? cfg.rotulos.coletarMais : cfg.rotulos.coletar;
            bBaixar.disabled = total === 0;
            bBaixar.textContent = `${cfg.rotulos.baixar} (${total})`;
            const bPdf = document.getElementById('btn-pdf');
            if (bPdf) { bPdf.disabled = total === 0; bPdf.textContent = `⬇ Baixar PDF (${total})`; }
            bLimpar.disabled = total === 0;

            if (total > 0) {
                const extra = cfg.usaAtuacao ? ` de ${contarAtuacoes()} atuação(ões)` : '';
                atualizarStatus(`Acumulado: ${total} registros${extra}.`);
            } else {
                atualizarStatus('');
            }
        }

        return { iniciar, continuar, baixar, pdf, limpar, render, rodando, obsoleta,
                 limparFlags: () => { store.removeItem(KEY_RODANDO); store.removeItem(KEY_TS); } };
    }

    // ── Navegação / paginação (comum aos dois relatórios) ───────────────────────

    function temProximaPagina() { return !!document.querySelector('a.arrowNextOn'); }

    function numeroPaginaAtual() {
        const b = document.querySelector('div#navigator b');
        return b ? parseInt(b.textContent.trim(), 10) : 1;
    }

    function totalRegistrosPagina() {
        const nav = document.querySelector('div#navigator div.navLeft');
        if (!nav) return null;
        const m = nav.textContent.match(/(\d+)\s+registro/);
        return m ? parseInt(m[1], 10) : null;
    }

    // ── Download genérico (dispara a partir de um clique do usuário) ─────────────

    function dataArquivo() { return new Date().toISOString().slice(0, 10); }

    function baixarBlob(blob, nome) {
        const url = URL.createObjectURL(blob);
        const a = document.createElement('a');
        a.href = url;
        a.download = nome;
        document.body.appendChild(a);
        a.click();
        setTimeout(() => { URL.revokeObjectURL(url); a.remove(); }, 2000);
    }

    // ── Geração e download do Excel ─────────────────────────────────────────────

    function gerarEbaixarExcel(dados, cfg) {
        const linhas = dados.map(cfg.linha);
        const wb = XLSX.utils.book_new();
        const ws = XLSX.utils.aoa_to_sheet([cfg.cabecalhos, ...linhas]);
        ws['!cols'] = cfg.larguras;
        XLSX.utils.book_append_sheet(wb, ws, 'Dados');

        const wbout = XLSX.write(wb, { bookType: 'xlsx', type: 'array' });
        const blob = new Blob([wbout], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
        baixarBlob(blob, `${cfg.nomeArquivo}_${dataArquivo()}.xlsx`);
    }

    // ── Geração e download do PDF (paisagem, painel + gráficos + tabela) ─────────

    // Paleta sóbria (tons de ardósia)
    const COR = {
        escura: [45, 55, 72],    // #2D3748
        media:  [90, 113, 132],  // barras
        clara:  [237, 242, 247], // #EDF2F7 (zebra / fundo de card)
        texto:  [45, 55, 72],
        suave:  [113, 128, 150], // textos secundários
        linha:  [203, 213, 224], // bordas
    };

    const DIA_MS = 86400000;

    function parseDataBR(str) {
        const m = /(\d{2})\/(\d{2})\/(\d{4})(?:\s+(\d{2}):(\d{2}))?/.exec(str || '');
        if (!m) return null;
        return new Date(+m[3], +m[2] - 1, +m[1], +(m[4] || 0), +(m[5] || 0)).getTime();
    }

    function diasDecorridos(str, now) {
        const ts = parseDataBR(str);
        if (ts == null) return '';
        return String(Math.max(0, Math.floor((now - ts) / DIA_MS)));
    }

    function acharMaisAntigo(dados, campoData) {
        let best = null;
        dados.forEach(d => {
            const ts = parseDataBR(d[campoData]);
            if (ts == null) return;
            if (!best || ts < best.ts) best = { ts, dataStr: (d[campoData] || '').trim(), registro: d };
        });
        return best;
    }

    function contarPorCampo(dados, campo, topN) {
        const mapa = new Map();
        dados.forEach(d => {
            const k = ((d[campo] || '').trim()) || '(vazio)';
            mapa.set(k, (mapa.get(k) || 0) + 1);
        });
        let arr = [...mapa.entries()].map(([label, valor]) => ({ label, valor })).sort((a, b) => b.valor - a.valor);
        if (arr.length > topN) {
            const resto = arr.slice(topN).reduce((s, i) => s + i.valor, 0);
            arr = arr.slice(0, topN);
            arr.push({ label: 'Outros', valor: resto });
        }
        return arr;
    }

    // Faixas de tempo de espera (com base nos dias decorridos até a emissão do relatório)
    function faixasDeTempo(dados, campoData, now) {
        let ate30 = 0, m30a90 = 0, mais90 = 0;
        dados.forEach(d => {
            const ts = parseDataBR(d[campoData]);
            if (ts == null) return;
            const dias = Math.floor((now - ts) / DIA_MS);
            if (dias > 90) mais90++;
            else if (dias > 30) m30a90++;
            else ate30++;
        });
        return [
            { label: 'Até 30 dias', valor: ate30, cor: [120, 140, 152] },
            { label: '31 a 90 dias', valor: m30a90, cor: [90, 113, 132] },
            { label: 'Mais de 90 dias', valor: mais90, cor: [150, 86, 86] },
        ];
    }

    // Card de destaque (KPI). subs: array de linhas secundárias.
    function desenharCard(doc, x, y, w, h, titulo, valor, subs) {
        doc.setDrawColor(...COR.linha); doc.setFillColor(...COR.clara); doc.setLineWidth(0.2);
        doc.roundedRect(x, y, w, h, 2, 2, 'FD');
        doc.setFont('helvetica', 'normal'); doc.setFontSize(8); doc.setTextColor(...COR.suave);
        doc.text(String(titulo).toUpperCase(), x + 5, y + 6.5);
        const grande = valor.length <= 13;
        doc.setFont('helvetica', 'bold'); doc.setFontSize(grande ? 19 : 14); doc.setTextColor(...COR.escura);
        doc.text(valor, x + 5, y + (grande ? 16 : 15));
        if (subs && subs.length) {
            doc.setFont('helvetica', 'normal'); doc.setFontSize(8); doc.setTextColor(...COR.suave);
            let yy = y + (grande ? 21 : 20);
            subs.forEach(s => {
                if (!s) return;
                doc.text(doc.splitTextToSize(String(s), w - 10)[0], x + 5, yy);
                yy += 4;
            });
        }
    }

    // Gráfico de barras horizontais. itens: [{label, valor, cor?}] na ordem de exibição.
    function desenharBarras(doc, x, y, w, h, titulo, itens) {
        doc.setFont('helvetica', 'bold'); doc.setFontSize(10); doc.setTextColor(...COR.escura);
        doc.text(titulo, x, y + 4);
        const topo = y + 9;
        const areaH = h - 9;
        if (!itens.length) return;
        const rotuloW = Math.min(64, w * 0.42);
        const valorW = 12;
        const barX = x + rotuloW;
        const barMaxW = Math.max(6, w - rotuloW - valorW);
        const maxVal = Math.max(...itens.map(i => i.valor)) || 1;
        const linhaH = Math.min(9, areaH / itens.length);
        const barH = Math.max(3, linhaH * 0.6);

        doc.setFontSize(8);
        itens.forEach((it, i) => {
            const meio = topo + i * linhaH + linhaH / 2;
            doc.setFont('helvetica', 'normal'); doc.setTextColor(...COR.texto);
            doc.text(doc.splitTextToSize(it.label, rotuloW - 3)[0], x, meio + 1.2);
            const bw = Math.max(0.6, (it.valor / maxVal) * barMaxW);
            doc.setFillColor(...(it.cor || COR.media));
            doc.roundedRect(barX, meio - barH / 2, bw, barH, 0.6, 0.6, 'F');
            doc.setFont('helvetica', 'bold'); doc.setTextColor(...COR.escura);
            doc.text(String(it.valor), barX + bw + 2, meio + 1.2);
        });
    }

    function gerarPDF(dados, cfg) {
        const ctor = (window.jspdf && window.jspdf.jsPDF) || window.jsPDF;
        if (!ctor) throw new Error('biblioteca jsPDF não carregada');
        const doc = new ctor({ orientation: 'landscape', unit: 'mm', format: 'a4' });
        if (typeof doc.autoTable !== 'function') throw new Error('plugin autoTable não carregado');

        const p = cfg.pdf;
        const now = Date.now();
        const pw = doc.internal.pageSize.getWidth();   // ~297 (paisagem)
        const ph = doc.internal.pageSize.getHeight();  // ~210
        const m = 12;
        const hoje = new Date().toLocaleDateString('pt-BR');

        // Ordena por data, da mais antiga para a mais nova (sem data vai para o fim)
        const ordenados = dados.slice().sort((a, b) => {
            const ta = parseDataBR(a[p.dataCampo]); const tb = parseDataBR(b[p.dataCampo]);
            return (ta == null ? Infinity : ta) - (tb == null ? Infinity : tb);
        });

        // Cabeçalho
        doc.setFont('helvetica', 'bold'); doc.setFontSize(16); doc.setTextColor(...COR.escura);
        doc.text(p.titulo, m, m + 2);
        doc.setFont('helvetica', 'normal'); doc.setFontSize(9); doc.setTextColor(...COR.suave);
        doc.text(`Relatório gerado em ${hoje}  •  ${dados.length} registro(s)`, m, m + 8);
        doc.setDrawColor(...COR.linha); doc.setLineWidth(0.3); doc.line(m, m + 11, pw - m, m + 11);

        // Painel: 2 KPIs (esquerda) + gráfico de faixas de tempo + gráfico de distribuição
        const topo = m + 16;
        const bandaH = 66;
        const gap = 6;

        // KPIs empilhados
        const kpiW = 72;
        desenharCard(doc, m, topo, kpiW, 28, p.atosTitulo, String(dados.length), ['atos pendentes']);

        const antigo = acharMaisAntigo(dados, p.dataCampo);
        let subsAntigo = ['Data não disponível'];
        let valAntigo = '—';
        if (antigo) {
            const reg = antigo.registro;
            const dias = Math.max(0, Math.floor((now - antigo.ts) / DIA_MS));
            valAntigo = antigo.dataStr;
            subsAntigo = [
                `Processo ${reg[p.processoCampo] || ''}`,
                reg[p.tipoCampo] || '',
                `${dias} dias em aberto`,
            ];
        }
        desenharCard(doc, m, topo + 28 + gap, kpiW, bandaH - 28 - gap, p.dataTitulo, valAntigo, subsAntigo);

        // Gráfico de faixas de tempo
        const agingX = m + kpiW + gap;
        const agingW = 74;
        desenharBarras(doc, agingX, topo, agingW, bandaH, p.agingTitulo, faixasDeTempo(dados, p.dataCampo, now));

        // Gráfico de distribuição (por agrupador / tipo de documento)
        const distX = agingX + agingW + gap;
        const distW = pw - m - distX;
        desenharBarras(doc, distX, topo, distW, bandaH, p.grupoTitulo, contarPorCampo(dados, p.grupoCampo, 12));

        // Tabela: colunas explícitas (larguras controladas => sem colunas sobrando), ordenada por data
        const colunas = p.colunas;
        const columnStyles = {};
        colunas.forEach((c, i) => { columnStyles[i] = { cellWidth: c.width }; });

        doc.autoTable({
            head: [colunas.map(c => c.header)],
            body: ordenados.map(d => colunas.map(c => String(c.get(d, { now }) ?? ''))),
            startY: topo + bandaH + 6,
            margin: { left: m, right: m, top: m + 4, bottom: 12 },
            theme: 'grid',
            tableWidth: 'wrap',
            styles: { font: 'helvetica', fontSize: 7.5, cellPadding: 1.6, textColor: COR.texto,
                      lineColor: COR.linha, lineWidth: 0.1, overflow: 'linebreak', valign: 'middle' },
            headStyles: { fillColor: COR.escura, textColor: [255, 255, 255], fontStyle: 'bold', fontSize: 8 },
            alternateRowStyles: { fillColor: COR.clara },
            columnStyles,
            didDrawPage: () => {
                const n = doc.internal.getNumberOfPages();
                doc.setFont('helvetica', 'normal'); doc.setFontSize(8); doc.setTextColor(...COR.suave);
                doc.text(`${p.titulo}  •  Página ${n}`, m, ph - 6);
                doc.text(hoje, pw - m, ph - 6, { align: 'right' });
            },
        });

        const blob = doc.output('blob');
        baixarBlob(blob, `${cfg.nomeArquivo}_${dataArquivo()}.pdf`);
    }

    // ── Interface ───────────────────────────────────────────────────────────────

    function atualizarStatus(msg) {
        const el = document.getElementById('exportar-status');
        if (el) el.textContent = msg;
    }

    function desabilitarBotoes(desabilitar) {
        ['btn-coletar', 'btn-baixar', 'btn-pdf', 'btn-limpar'].forEach(id => {
            const b = document.getElementById(id);
            if (b) b.disabled = desabilitar;
        });
    }

    GM_addStyle(`
        .projudi-btn {
            color: white; border: 1px solid #145214; padding: 3px 10px; cursor: pointer;
            font-size: 0.85em; font-family: Verdana, Arial, sans-serif; font-weight: bold;
            margin-left: 6px; border-radius: 3px;
        }
        #btn-coletar { background-color: #1e6b1e; border-color: #145214; }
        #btn-baixar  { background-color: #b8860b; border-color: #8a6508; }
        #btn-pdf     { background-color: #34556b; border-color: #26404f; }
        #btn-limpar  { background-color: #8a3b3b; border-color: #6e2f2f; }
        .projudi-btn:disabled { background-color: #999; border-color: #777; cursor: not-allowed; }
        #exportar-status { font-size: 0.8em; color: #555; margin-left: 8px; font-family: Verdana, Arial, sans-serif; }
    `);

    // Detecta qual relatório está na tela pelo cabeçalho da resultTable; se não houver
    // tabela reconhecível, tenta pela URL (útil na tela de busca da página de juntadas).
    function detectarConfig() {
        const thead = document.querySelector('table.resultTable thead');
        const cab = thead ? thead.textContent : '';
        if (CFG_JUNTADAS.detecta(cab)) return CFG_JUNTADAS;
        if (CFG_RETORNO.detecta(cab)) return CFG_RETORNO;
        if (CFG_CONCLUSOES.detecta(cab)) return CFG_CONCLUSOES;
        if (/analisarJuntada\.do/i.test(location.pathname + location.search)) return CFG_JUNTADAS;
        return null; // sem tabela de resultados reconhecível (página de conclusao.do sem resultados)
    }

    function injetarBotoes() {
        const buttonBar = document.querySelector('table.buttonBar td.buttons');
        if (!buttonBar) return;

        // Descobre o relatório atual; se não houver tabela reconhecível, assume Conclusões
        // (mas ainda respeita uma coleta de Retorno em andamento, retomada após reload).
        let cfg = detectarConfig();
        if (!cfg) {
            // Página de conclusao.do sem resultados: respeita uma coleta em andamento;
            // senão, assume o relatório de Conclusões.
            const emAndamento = [CFG_RETORNO, CFG_JUNTADAS, CFG_CONCLUSOES]
                .find(c => store.getItem(c.prefixo + 'rodando') === '1');
            cfg = emAndamento || CFG_CONCLUSOES;
        }
        const coletor = criarColetor(cfg);

        const mk = (id, title, onclick, texto) => {
            const b = document.createElement('button');
            b.id = id; b.type = 'button'; b.className = 'projudi-btn';
            b.title = title; b.onclick = onclick;
            if (texto) b.textContent = texto;
            return b;
        };

        buttonBar.appendChild(mk('btn-coletar', 'Percorre todas as páginas e acrescenta aos dados já coletados', () => coletor.iniciar()));
        buttonBar.appendChild(mk('btn-baixar', 'Junta tudo o que foi coletado e baixa a planilha Excel', () => coletor.baixar()));
        if (cfg.pdf) {
            buttonBar.appendChild(mk('btn-pdf', 'Gera um PDF (paisagem) com painel, gráficos e a tabela completa', () => coletor.pdf()));
        }
        buttonBar.appendChild(mk('btn-limpar', 'Apaga os dados acumulados deste relatório', () => coletor.limpar(), 'Limpar'));

        const status = document.createElement('span');
        status.id = 'exportar-status';
        buttonBar.appendChild(status);

        if (coletor.rodando() && !coletor.obsoleta()) {
            coletor.continuar(); // retoma após o reload da paginação
        } else {
            coletor.limparFlags(); // descarta flag de execução presa, mantendo os dados
            coletor.render();
        }
    }

    if (document.readyState === 'loading') {
        document.addEventListener('DOMContentLoaded', injetarBotoes);
    } else {
        injetarBotoes();
    }
})();
