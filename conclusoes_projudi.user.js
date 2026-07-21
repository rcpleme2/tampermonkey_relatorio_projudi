// ==UserScript==
// @name         Exportar Conclusões Projudi para Excel
// @namespace    https://projudi2.tjpr.jus.br/
// @version      8.3
// @description  Coleta conclusões (remessa), retorno de conclusos ou juntadas, acumula e exporta em Excel ou PDF (retrato, resumo com KPIs e gráficos, prioritários destacados)
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

    // Competência = o nome completo da Atuação (ex.: "Vara de Família de Curitiba")
    function competenciaDe(atuacao) {
        return (atuacao || '').trim();
    }

    // Processo prioritário: o número recebe uma classe de cor extra (attentionBlue,
    // attentionPurple, attentionRed, ...); o normal fica apenas com "attention" (preto).
    function emPrioritario(em) {
        if (!em) return false;
        return (em.className || '').split(/\s+/)
            .some(c => /^attention[A-Z]/.test(c) && c !== 'attentionBold');
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
        cabecalhos: ['Atuação', 'Dt. Remessa', 'Processo', 'Classe', 'Seq.', 'Tipo de Conclusão', 'Privativa', 'Responsável', 'Pré-análise', 'Agrupador', 'Prioritário'],
        larguras: [{ wch: 24 }, { wch: 18 }, { wch: 26 }, { wch: 18 }, { wch: 6 }, { wch: 30 }, { wch: 10 }, { wch: 30 }, { wch: 14 }, { wch: 20 }, { wch: 11 }],
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
                prioritario: emPrioritario(tds[2].querySelector('em')),
            };
        },
        linha: (d) => [d.atuacao, d.dtRemessa, d.processo, d.classe, d.seq, d.tipoConclusao, d.privativa, d.responsavel, d.preAnalise, d.agrupador, d.prioritario ? 'Sim' : 'Não'],
    };

    const CFG_RETORNO = {
        prefixo: 'projudi_retorno_',
        detecta: (cab) => /retorno/i.test(cab),
        minTds: 9,                              // linhas têm 10 tds (0 = semáforo)
        usaAtuacao: false,
        nomeArquivo: 'retorno_conclusos_projudi',
        rotulos: { coletar: 'Extrair Retorno', coletarMais: 'Extrair mais (Retorno)', baixar: '⬇ Baixar Retorno' },
        // Colunas pedidas: Processo - Classe - Dt. Retorno - Tipo de conclusão - Responsável - Agrupador
        cabecalhos: ['Processo', 'Classe', 'Dt. Retorno', 'Tipo de Conclusão', 'Responsável', 'Agrupador', 'Prioritário'],
        larguras: [{ wch: 26 }, { wch: 18 }, { wch: 18 }, { wch: 30 }, { wch: 30 }, { wch: 20 }, { wch: 11 }],
        extrai: (tds, atuacao) => {
            const pc = processoEclasse(tds[3]);
            return {
                processo: pc.processo,
                classe: pc.classe,
                dtRetorno: textoCelula(tds[2]),
                tipoConclusao: textoCelula(tds[5]),
                responsavel: textoCelula(tds[7]),
                agrupador: textoCelula(tds[9]),
                prioritario: emPrioritario(tds[3].querySelector('em')),
                atuacao: atuacao || '',
                competencia: competenciaDe(atuacao),
            };
        },
        linha: (d) => [d.processo, d.classe, d.dtRetorno, d.tipoConclusao, d.responsavel, d.agrupador, d.prioritario ? 'Sim' : 'Não'],
        pdf: {
            titulo: 'Retorno de Processos Conclusos',
            atosTitulo: 'Retornos de conclusão pendentes',
            agingTitulo: 'Retornos por tempo de espera',
            tabelaTitulo: 'Tabela discriminada dos retornos de conclusão pendentes',
            dataCampo: 'dtRetorno',
            dataTitulo: 'Retorno de conclusão mais antigo',
            processoCampo: 'processo',
            tipoCampo: 'tipoConclusao',
            // mediaLabel ausente => o KPI de média por dia não é exibido no retorno
            distribuicoes: [
                { titulo: 'Processos por Agrupador', campo: 'agrupador', topN: 12 },
                { titulo: 'Retornos pendentes por magistrado', campo: 'responsavel', topN: 12 },
            ],
            // Colunas (retrato): Dt. Retorno logo antes de Dias
            colunas: [
                { header: 'Processo', width: 30, get: (d) => d.processo },
                { header: 'Classe', width: 20, get: (d) => d.classe },
                { header: 'Tipo de Conclusão', width: 34, get: (d) => d.tipoConclusao },
                { header: 'Responsável', width: 34, get: (d) => d.responsavel },
                { header: 'Agrupador', width: 26, get: (d) => d.agrupador },
                { header: 'Dt. Retorno', width: 24, get: (d) => d.dtRetorno },
                { header: 'Dias', width: 12, get: (d, ctx) => diasDecorridos(d.dtRetorno, ctx.now) },
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
        cabecalhos: ['Processo', 'Tipo de Documento', 'Especificação do Documento', 'Data de Envio', 'Juntado por', 'Função', 'Prioritário'],
        larguras: [{ wch: 26 }, { wch: 40 }, { wch: 52 }, { wch: 18 }, { wch: 28 }, { wch: 18 }, { wch: 11 }],
        extrai: (tds, atuacao) => {
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
                funcao: (bFunc ? norm(bFunc.textContent) : '') || 'Sistema',
                prioritario: emPrioritario(emProc),
                atuacao: atuacao || '',
                competencia: competenciaDe(atuacao),
            };
        },
        linha: (d) => [d.processo, d.tipoDocumento, d.especificacao, d.dataEnvio, d.juntadoPor, d.funcao || '', d.prioritario ? 'Sim' : 'Não'],
        pdf: {
            titulo: 'Juntadas',
            atosTitulo: 'Juntadas pendentes',
            agingTitulo: 'Juntadas por tempo de espera',
            tabelaTitulo: 'Tabela discriminada das juntadas pendentes',
            dataCampo: 'dataEnvio',
            dataTitulo: 'Juntada pendente mais antiga',
            processoCampo: 'processo',
            tipoCampo: 'tipoDocumento',
            mediaLabel: 'juntadas / dia',
            distribuicoes: [
                { titulo: 'Pendências por Função', campo: 'funcao', topN: 10 },
                { titulo: 'Juntadas pendentes por pessoa', campo: 'juntadoPor', topN: 12 },
                // Largura total (span 2) para caber o nome completo do tipo de documento
                { titulo: 'Processos por Tipo de Documento', campo: 'tipoDocumento', topN: 14, span: 2, limpar: (s) => s.replace(/^juntada de\s+/i, '') },
            ],
            // Colunas (retrato): Data de Envio logo antes de Dias
            colunas: [
                { header: 'Processo', width: 30, get: (d) => d.processo },
                { header: 'Tipo de Documento', width: 30, get: (d) => d.tipoDocumento },
                { header: 'Especificação do Documento', width: 46, get: (d) => d.especificacao },
                { header: 'Juntado por', width: 30, get: (d) => d.juntadoPor },
                { header: 'Função', width: 16, get: (d) => d.funcao || '' },
                { header: 'Data de Envio', width: 22, get: (d) => d.dataEnvio },
                { header: 'Dias', width: 12, get: (d, ctx) => diasDecorridos(d.dataEnvio, ctx.now) },
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
            // Sempre lê a Atuação (usada para a competência nos relatórios de retorno/juntadas)
            const atuacao = lerAtuacao();
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

    function contarPorCampo(dados, campo, topN, limpar) {
        const mapa = new Map();
        dados.forEach(d => {
            let k = (d[campo] || '').trim();
            if (limpar) k = limpar(k).trim();
            k = k || '(vazio)';
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

    const COR_PRIORITARIO = [180, 60, 60];  // realce dos prioritários

    function contarPrioritarios(dados) {
        return dados.reduce((n, d) => n + (d.prioritario ? 1 : 0), 0);
    }

    // Média de registros por dia (com base nos dias distintos presentes nas datas)
    function mediaPorDia(dados, campoData) {
        const dias = new Set();
        let comData = 0;
        dados.forEach(d => {
            const m = /(\d{2})\/(\d{2})\/(\d{4})/.exec(d[campoData] || '');
            if (!m) return;
            comData++;
            dias.add(m[0]);
        });
        if (!dias.size) return 0;
        return comData / dias.size;
    }

    // Faixas de tempo de espera, separando prioritários de normais.
    function faixasPorPrioridade(dados, campoData, now) {
        const b = [
            { label: 'Até 30 dias', prioritarios: 0, normais: 0 },
            { label: '31 a 90 dias', prioritarios: 0, normais: 0 },
            { label: 'Mais de 90 dias', prioritarios: 0, normais: 0 },
        ];
        dados.forEach(d => {
            const ts = parseDataBR(d[campoData]);
            if (ts == null) return;
            const dias = Math.floor((now - ts) / DIA_MS);
            const idx = dias > 90 ? 2 : (dias > 30 ? 1 : 0);
            if (d.prioritario) b[idx].prioritarios++; else b[idx].normais++;
        });
        return b;
    }

    // Card de destaque (KPI). subs: array de linhas secundárias. Se central=true, o
    // conteúdo é centralizado horizontal e verticalmente dentro do card.
    function desenharCard(doc, x, y, w, h, titulo, valor, subs, central) {
        doc.setDrawColor(...COR.linha); doc.setFillColor(...COR.clara); doc.setLineWidth(0.2);
        doc.roundedRect(x, y, w, h, 2, 2, 'FD');
        subs = (subs || []).filter(Boolean);

        if (central) {
            const cx = x + w / 2;
            // altura total do bloco: título (4) + valor (7) + subs (4 cada)
            const blocoH = 4 + 8 + subs.length * 4.2;
            let yy = y + (h - blocoH) / 2 + 4;
            doc.setFont('helvetica', 'normal'); doc.setFontSize(8); doc.setTextColor(...COR.suave);
            doc.text(String(titulo).toUpperCase(), cx, yy, { align: 'center' }); yy += 7;
            doc.setFont('helvetica', 'bold'); doc.setFontSize(valor.length <= 26 ? 14 : 11); doc.setTextColor(...COR.escura);
            doc.text(valor, cx, yy, { align: 'center' }); yy += 5.5;
            doc.setFont('helvetica', 'normal'); doc.setFontSize(8); doc.setTextColor(...COR.suave);
            subs.forEach(s => { doc.text(doc.splitTextToSize(String(s), w - 8)[0], cx, yy, { align: 'center' }); yy += 4.2; });
            return;
        }

        doc.setFont('helvetica', 'normal'); doc.setFontSize(8); doc.setTextColor(...COR.suave);
        doc.text(String(titulo).toUpperCase(), x + 5, y + 6.5);
        const grande = valor.length <= 13;
        doc.setFont('helvetica', 'bold'); doc.setFontSize(grande ? 19 : 14); doc.setTextColor(...COR.escura);
        doc.text(valor, x + 5, y + (grande ? 16 : 15));
        if (subs.length) {
            doc.setFont('helvetica', 'normal'); doc.setFontSize(8); doc.setTextColor(...COR.suave);
            let yy = y + (grande ? 21 : 20);
            subs.forEach(s => { doc.text(doc.splitTextToSize(String(s), w - 10)[0], x + 5, yy); yy += 4; });
        }
    }

    // Gráfico de barras horizontais. itens: [{label, valor, cor?}] na ordem de exibição.
    function desenharBarras(doc, x, y, w, h, titulo, itens) {
        doc.setFont('helvetica', 'bold'); doc.setFontSize(10); doc.setTextColor(...COR.escura);
        doc.text(titulo, x, y + 4);
        const topo = y + 9;
        const areaH = h - 9;
        if (!itens.length) return;
        // Rótulo ocupa metade da largura (útil sobretudo no gráfico de largura total,
        // para caber o nome completo do tipo de documento em vez de truncar em "PETIÇÃO DE")
        const rotuloW = Math.min(w * 0.5, 120);
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

    // Barras agrupadas por faixa: duas sub-barras (prioritários x normais) por linha, com legenda.
    function desenharBarrasFaixas(doc, x, y, w, h, titulo, faixas) {
        // Título (linha 1) e legenda (linha 2) — separados para não se sobreporem
        doc.setFont('helvetica', 'bold'); doc.setFontSize(9.5); doc.setTextColor(...COR.escura);
        doc.text(doc.splitTextToSize(titulo, w)[0], x, y + 4);
        const legY = y + 9.5;
        doc.setFillColor(...COR_PRIORITARIO); doc.rect(x, legY - 2.4, 3, 3, 'F');
        doc.setFont('helvetica', 'normal'); doc.setFontSize(7); doc.setTextColor(...COR.texto);
        doc.text('Prioritários', x + 4, legY);
        const q2 = x + 28;
        doc.setFillColor(...COR.media); doc.rect(q2, legY - 2.4, 3, 3, 'F');
        doc.text('Normais', q2 + 4, legY);

        const topo = y + 14;
        const areaH = h - 14;
        const rotuloW = Math.min(34, w * 0.32);
        const valorW = 10;
        const barX = x + rotuloW;
        const barMaxW = Math.max(6, w - rotuloW - valorW);
        const maxVal = Math.max(1, ...faixas.map(f => Math.max(f.prioritarios, f.normais)));
        const linhaH = areaH / faixas.length;
        const subH = Math.max(2.4, linhaH * 0.26);

        doc.setFontSize(7.5);
        faixas.forEach((f, i) => {
            const base = topo + i * linhaH;
            const meio = base + linhaH / 2;
            doc.setFont('helvetica', 'normal'); doc.setTextColor(...COR.texto);
            doc.text(doc.splitTextToSize(f.label, rotuloW - 3)[0], x, meio + 1);
            // prioritários (acima)
            const wp = Math.max(0.5, (f.prioritarios / maxVal) * barMaxW);
            doc.setFillColor(...COR_PRIORITARIO);
            doc.roundedRect(barX, meio - subH - 1, wp, subH, 0.5, 0.5, 'F');
            doc.setFont('helvetica', 'bold'); doc.setTextColor(...COR.escura);
            doc.text(String(f.prioritarios), barX + wp + 1.5, meio - 1.4);
            // normais (abaixo)
            const wn = Math.max(0.5, (f.normais / maxVal) * barMaxW);
            doc.setFillColor(...COR.media);
            doc.roundedRect(barX, meio + 1, wn, subH, 0.5, 0.5, 'F');
            doc.text(String(f.normais), barX + wn + 1.5, meio + subH + 0.6);
        });
    }

    function desenharRodape(doc, titulo, quando, pw, ph, m) {
        doc.setFont('helvetica', 'normal'); doc.setFontSize(8); doc.setTextColor(...COR.suave);
        doc.text(`${titulo}  •  Página ${doc.internal.getNumberOfPages()}`, m, ph - 6);
        doc.text(quando, pw - m, ph - 6, { align: 'right' });
    }

    function gerarPDF(dados, cfg) {
        const ctor = (window.jspdf && window.jspdf.jsPDF) || window.jsPDF;
        if (!ctor) throw new Error('biblioteca jsPDF não carregada');
        const doc = new ctor({ orientation: 'portrait', unit: 'mm', format: 'a4' });
        if (typeof doc.autoTable !== 'function') throw new Error('plugin autoTable não carregado');

        const p = cfg.pdf;
        const now = Date.now();
        const agora = new Date();
        const pw = doc.internal.pageSize.getWidth();   // ~210 (retrato)
        const ph = doc.internal.pageSize.getHeight();  // ~297
        const m = 12;
        const uw = pw - 2 * m;                          // largura útil
        const hoje = agora.toLocaleDateString('pt-BR');
        const hora = agora.toLocaleTimeString('pt-BR', { hour: '2-digit', minute: '2-digit' });
        const carimbo = `${hoje} ${hora}`;

        // Ordena por data, da mais antiga para a mais nova (sem data vai para o fim)
        const ordenados = dados.slice().sort((a, b) => {
            const ta = parseDataBR(a[p.dataCampo]); const tb = parseDataBR(b[p.dataCampo]);
            return (ta == null ? Infinity : ta) - (tb == null ? Infinity : tb);
        });

        // Desenha uma página de resumo (KPIs + gráficos) para um subconjunto de dados.
        const gap = 6;
        function desenharPaginaResumo(sub, contexto, primeira) {
            if (!primeira) doc.addPage();

            // Cabeçalho: título, competência em destaque (com quebra), e a data em linha própria
            let hy = m + 2;
            doc.setFont('helvetica', 'bold'); doc.setFontSize(16); doc.setTextColor(...COR.escura);
            doc.text(p.titulo, m, hy);
            hy += 8;
            if (contexto.competencia) {
                doc.setFont('helvetica', 'bold'); doc.setFontSize(11.5); doc.setTextColor(...COR.media);
                const linhas = doc.splitTextToSize('Competência: ' + contexto.competencia, uw);
                doc.text(linhas, m, hy);
                hy += linhas.length * 5.2 + 1.5;
            } else if (contexto.rotulo) {
                doc.setFont('helvetica', 'bold'); doc.setFontSize(11.5); doc.setTextColor(...COR.media);
                doc.text(contexto.rotulo, m, hy);
                hy += 7;
            }
            doc.setFont('helvetica', 'normal'); doc.setFontSize(9); doc.setTextColor(...COR.suave);
            doc.text(`Extraído em ${hoje} às ${hora}  •  ${sub.length} registro(s)`, m, hy);
            hy += 3;
            doc.setDrawColor(...COR.linha); doc.setLineWidth(0.3); doc.line(m, hy, pw - m, hy);

            // KPIs numéricos (média por dia só quando o relatório define mediaLabel)
            const kY = hy + 5;
            const prio = contarPrioritarios(sub);
            const kpis = [
                { titulo: p.atosTitulo, valor: String(sub.length), subs: [] },
                { titulo: 'Prioritários pendentes', valor: String(prio), subs: [`${sub.length ? Math.round(prio / sub.length * 100) : 0}% do total`] },
            ];
            if (p.mediaLabel) {
                const media = mediaPorDia(sub, p.dataCampo);
                kpis.push({ titulo: 'Média por dia', valor: media ? media.toLocaleString('pt-BR', { maximumFractionDigits: 1 }) : '—', subs: [p.mediaLabel] });
            }
            const kW = (uw - (kpis.length - 1) * gap) / kpis.length;
            kpis.forEach((k, i) => desenharCard(doc, m + i * (kW + gap), kY, kW, 28, k.titulo, k.valor, k.subs));

            // KPI do mais atrasado (card largo, texto centralizado)
            const aY = kY + 28 + gap;
            const antigo = acharMaisAntigo(sub, p.dataCampo);
            let subsAntigo = ['Data não disponível'];
            let valAntigo = '—';
            if (antigo) {
                const reg = antigo.registro;
                const dias = Math.max(0, Math.floor((now - antigo.ts) / DIA_MS));
                valAntigo = `${antigo.dataStr}  (${dias} dias em aberto)`;
                subsAntigo = [
                    `Processo ${reg[p.processoCampo] || ''}${reg.prioritario ? '  — PRIORITÁRIO' : ''}`,
                    reg[p.tipoCampo] || '',
                ];
            }
            desenharCard(doc, m, aY, uw, 28, p.dataTitulo, valAntigo, subsAntigo, true);

            // Gráficos (grade de 2 colunas; um gráfico pode ocupar as 2 colunas com span:2)
            const charts = [
                { tipo: 'faixas', span: 1, titulo: p.agingTitulo, faixas: faixasPorPrioridade(sub, p.dataCampo, now) },
                ...p.distribuicoes.map(g => ({ tipo: 'barras', span: g.span || 1, titulo: g.titulo, itens: contarPorCampo(sub, g.campo, g.topN, g.limpar) })),
            ];
            let col = 0, row = 0;
            charts.forEach(c => {
                if (c.span === 2) {
                    if (col !== 0) { row++; col = 0; }
                    c.pos = { row, col: 0, span: 2 }; row++; col = 0;
                } else {
                    c.pos = { row, col, span: 1 }; col++;
                    if (col >= 2) { col = 0; row++; }
                }
            });
            const nLinhas = Math.max(...charts.map(c => c.pos.row)) + 1;
            const gY0 = aY + 28 + gap + 2;
            const colW = (uw - gap) / 2;
            const chartH = Math.min(96, (ph - m - gY0 - (nLinhas - 1) * 10) / nLinhas);
            charts.forEach(c => {
                const cx = m + c.pos.col * (colW + gap);
                const cy = gY0 + c.pos.row * (chartH + 10);
                const cw = c.pos.span === 2 ? uw : colW;
                if (c.tipo === 'faixas') desenharBarrasFaixas(doc, cx, cy, cw, chartH, c.titulo, c.faixas);
                else desenharBarras(doc, cx, cy, cw, chartH, c.titulo, c.itens);
            });

            desenharRodape(doc, p.titulo, carimbo, pw, ph, m);
        }

        // ═══ RESUMO GERAL ═══
        desenharPaginaResumo(dados, { rotulo: 'Resumo geral' }, true);

        // ═══ RESUMO POR COMPETÊNCIA (quando há mais de uma) ═══
        const porComp = new Map();
        dados.forEach(d => {
            const c = (d.competencia || '').trim();
            if (!c) return;
            if (!porComp.has(c)) porComp.set(c, []);
            porComp.get(c).push(d);
        });
        if (porComp.size > 1) {
            [...porComp.entries()]
                .sort((a, b) => b[1].length - a[1].length)
                .forEach(([comp, sub]) => desenharPaginaResumo(sub, { competencia: comp }, false));
        }

        // ═══ TABELA DISCRIMINADA ═══
        doc.addPage();
        doc.setFont('helvetica', 'bold'); doc.setFontSize(12); doc.setTextColor(...COR.escura);
        doc.text(p.tabelaTitulo || 'Tabela discriminada', m, m + 3);
        const tabInicioY = m + 8;
        const colunas = p.colunas;
        const columnStyles = {};
        colunas.forEach((c, i) => { columnStyles['k' + i] = { cellWidth: c.width }; });
        const idxProcesso = colunas.findIndex(c => /processo/i.test(c.header));

        doc.autoTable({
            columns: colunas.map((c, i) => ({ header: c.header, dataKey: 'k' + i })),
            body: ordenados.map(d => {
                const o = {};
                colunas.forEach((c, i) => { o['k' + i] = String(c.get(d, { now }) ?? ''); });
                return o;
            }),
            startY: tabInicioY,
            margin: { left: m, right: m, top: m, bottom: 14 },
            theme: 'grid',
            styles: { font: 'helvetica', fontSize: 7.5, cellPadding: 1.6, textColor: COR.texto,
                      lineColor: COR.linha, lineWidth: 0.1, overflow: 'linebreak', valign: 'middle' },
            headStyles: { fillColor: COR.escura, textColor: [255, 255, 255], fontStyle: 'bold', fontSize: 8 },
            alternateRowStyles: { fillColor: COR.clara },
            columnStyles,
            // Realça o número do processo dos prioritários
            didParseCell: (data) => {
                if (data.section === 'body' && data.column.index === idxProcesso && ordenados[data.row.index] && ordenados[data.row.index].prioritario) {
                    data.cell.styles.textColor = COR_PRIORITARIO;
                    data.cell.styles.fontStyle = 'bold';
                }
            },
            didDrawPage: () => desenharRodape(doc, p.titulo, carimbo, pw, ph, m),
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
