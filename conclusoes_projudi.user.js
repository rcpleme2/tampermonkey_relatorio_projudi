// ==UserScript==
// @name         Exportar Conclusões Projudi para Excel
// @namespace    https://projudi2.tjpr.jus.br/
// @version      12.0
// @description  Coleta conclusões/retorno/juntadas/tempo médio, exporta Excel ou PDF, e automatiza a extração conjunta a partir da página inicial
// @author       rcpleme2
// @match        https://projudi2.tjpr.jus.br/projudi/*
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

    // Texto até o primeiro <br> (ex.: pegar só "DECISÃO" numa célula que tem juiz/pré-análise abaixo)
    function textoAteBr(td) {
        if (!td) return '';
        let s = '';
        for (const n of td.childNodes) {
            if (n.nodeType === 1 && n.tagName === 'BR') break;
            s += n.textContent;
        }
        return norm(s);
    }

    // Extrai a primeira data dd/mm/aaaa de um texto (ex.: "03/07/2026 04680494956.est")
    function dataDeTexto(td) {
        const m = /(\d{2}\/\d{2}\/\d{4})/.exec(td ? td.textContent : '');
        return m ? m[1] : '';
    }

    // Competência = o nome completo da Atuação (ex.: "Vara de Família de Curitiba")
    function competenciaDe(atuacao) {
        return (atuacao || '').trim();
    }

    // Processo prioritário: o número recebe uma classe de cor extra (attentionBlue,
    // attentionPurple, attentionRed, ...); o normal fica apenas com "attention" (preto).
    function emPrioritario(em) {
        if (!em) return false;
        // Verifica classes de cor (attentionRed, attentionBlue, attentionPurple …)
        if ((em.className || '').split(/\s+/).some(c => /^attention[A-Z]/.test(c) && c !== 'attentionBold')) return true;
        // Fallback: cor computada — vermelho indica processo prioritário (relatório tempoMédio)
        try {
            const cor = window.getComputedStyle(em).color;
            const m = /rgb\((\d+),\s*(\d+),\s*(\d+)\)/.exec(cor);
            if (m && +m[1] > 150 && +m[2] < 100 && +m[3] < 100) return true;
        } catch (e) {}
        return false;
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

    // Relatório de Estatística de Conclusão (Analisadas, Analítico) — tempo de cumprimento.
    // Colunas da tabela: [0]expandir [1]Processo [2]Dt.Envio [3]Dt.Análise
    //                    [4]Dt.Análise Cartório [5]Tipo de conclusão [6]Classe Processual
    const CFG_TEMPOMEDIO = {
        prefixo: 'projudi_tempomedio_',
        detecta: (cab) => /an[áa]lise\s+cart[óo]rio/i.test(cab),
        minTds: 7,
        usaAtuacao: false,
        nomeArquivo: 'tempo_medio_projudi',
        rotulos: { coletar: 'Extrair Tempo Médio', coletarMais: 'Extrair mais (Tempo Médio)', baixar: '⬇ Baixar Tempo Médio' },
        cabecalhos: ['Processo', 'Dt. Análise', 'Dt. Análise Cartório', 'Tipo de Conclusão', 'Classe Processual', 'Dias p/ Cumprimento', 'Prioritário'],
        larguras: [{ wch: 26 }, { wch: 16 }, { wch: 20 }, { wch: 26 }, { wch: 40 }, { wch: 18 }, { wch: 11 }],
        extrai: (tds, atuacao) => {
            const emProc = tds[1].querySelector('em');
            const processo = emProc ? emProc.textContent.trim() : textoCelula(tds[1]);
            const dtAnalise = dataDeTexto(tds[3]);
            const dtCartorio = dataDeTexto(tds[4]);
            const classeCompleta = textoAteBr(tds[6]) || textoCelula(tds[6]);
            const classe = classeCompleta.split(' (')[0].trim(); // classe sem o "(Assunto Principal)"
            // Dias para cumprimento = Dt. Análise Cartório - Dt. Análise
            const tA = parseDataBR(dtAnalise), tC = parseDataBR(dtCartorio);
            const dias = (tA != null && tC != null) ? Math.max(0, Math.round((tC - tA) / DIA_MS)) : null;
            return {
                processo,
                dtAnalise,
                dtCartorio,
                tipoConclusao: textoAteBr(tds[5]),
                classe,
                dias,
                prioritario: emPrioritario(emProc),
                atuacao: atuacao || '',
                competencia: competenciaDe(atuacao),
            };
        },
        linha: (d) => [d.processo, d.dtAnalise, d.dtCartorio, d.tipoConclusao, d.classe,
                       (d.dias == null ? '' : String(d.dias)), d.prioritario ? 'Sim' : 'Não'],
        pdfCustom: (dados) => gerarPDFTempoMedio(dados),
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
            const atuacao = lerAtuacao();
            const linhas = document.querySelectorAll('table.resultTable tbody tr');
            const dados = [];
            linhas.forEach(tr => {
                const tds = tr.querySelectorAll('td');
                if (tds.length < cfg.minTds) return;
                dados.push(cfg.extrai(tds, atuacao));
            });
            console.log(`[Projudi] coletarPaginaAtual — ${linhas.length} linhas encontradas, ${dados.length} extraídas (minTds=${cfg.minTds})`);
            return dados;
        }

        function iniciar() {
            // Antes de iniciar, garante que a página exibe 500 registros por vez para
            // minimizar o número de recargas. Quando o valor do select muda, o Projudi
            // recarrega a página; o estado KEY_RODANDO já estará salvo e continuar()
            // será chamado automaticamente ao reiniciar (bloco rodando && !obsoleta).
            const sel = document.querySelector('select[name="estatisticaPageSizeOptions"]');
            if (sel && sel.value !== '500') {
                console.log(`[Projudi] alterando pageSize de ${sel.value} para 500 — aguardando reload`);
                store.setItem(KEY_RODANDO, '1');
                marcarAtividade();
                sel.value = '500';
                sel.dispatchEvent(new Event('change', { bubbles: true }));
                return;
            }
            console.log('[Projudi] iniciar() — pageSize OK, iniciando continuar()');
            store.setItem(KEY_RODANDO, '1');
            marcarAtividade();
            continuar();
        }

        function continuar() {
            desabilitarBotoes(true);
            marcarAtividade();
            console.log('[Projudi] continuar() — coletando página atual');

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
                avancarAutomacao(cfg); // se a automação estiver ativa, segue para o próximo passo
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
                if (cfg.pdfCustom) cfg.pdfCustom(dados); else gerarPDF(dados, cfg);
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

    // ── Geração e download do PDF (retrato, KPIs + gráficos + tabela) ────────────

    // Paleta de acento (validada para acessibilidade/colorblind-safe). Uso semântico
    // fixo: azul = métrica principal, vermelho = prioritário (sempre), âmbar = atenção/
    // tempo de espera intermediário, aqua = positivo/secundário. Cor nunca é o único
    // sinal — sempre acompanhada de rótulo ou legenda.
    const COR = {
        tinta:    [17, 17, 17],     // títulos fortes
        tintaSec: [82, 81, 78],     // texto secundário
        muted:    [137, 135, 129],  // eixos / rótulos em maiúsculas
        grade:    [225, 224, 217],  // hairline / bordas
        base:     [195, 194, 183],  // linhas de base
        cartao:   [244, 247, 251],  // fundo de card
        azul:     [42, 120, 214],   // métrica principal / série "normal"
        aqua:     [27, 175, 122],   // secundário / positivo
        ambar:    [250, 178, 25],   // atenção / faixa intermediária
        vermelho: [208, 59, 59],    // PRIORITÁRIO / crítico
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

    const COR_PRIORITARIO = COR.vermelho;  // realce dos prioritários (mesma cor em todo o relatório)

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

    // Título de seção com régua de acento (usado por gráficos e blocos da tabela).
    function tituloSecao(doc, x, y, w, texto, acento) {
        acento = acento || COR.azul;
        doc.setFont('helvetica', 'bold'); doc.setFontSize(10); doc.setTextColor(...COR.tinta);
        doc.text(texto, x, y);
        doc.setDrawColor(...acento); doc.setLineWidth(0.7);
        doc.line(x, y + 1.6, x + Math.min(w, doc.getTextWidth(texto) + 2), y + 1.6);
    }

    // Card de destaque (KPI), com uma barra de acento colorida à esquerda indicando o
    // papel semântico do dado (azul=principal, vermelho=prioritário, âmbar=atenção,
    // aqua=secundário). subs: array de linhas secundárias. Se central=true, o conteúdo
    // é centralizado horizontal e verticalmente dentro do card.
    function desenharCard(doc, x, y, w, h, titulo, valor, subs, central, acento) {
        acento = acento || COR.azul;
        doc.setDrawColor(...COR.grade); doc.setFillColor(...COR.cartao); doc.setLineWidth(0.2);
        doc.roundedRect(x, y, w, h, 2, 2, 'FD');
        doc.setFillColor(...acento);
        doc.roundedRect(x, y, 1.8, h, 0.9, 0.9, 'F');
        subs = (subs || []).filter(Boolean);

        if (central) {
            const cx = x + w / 2;
            // altura total do bloco: título (4) + valor (7) + subs (4 cada)
            const blocoH = 4 + 8 + subs.length * 4.2;
            let yy = y + (h - blocoH) / 2 + 4;
            doc.setFont('helvetica', 'bold'); doc.setFontSize(7.5); doc.setTextColor(...COR.muted);
            doc.text(String(titulo).toUpperCase(), cx, yy, { align: 'center' }); yy += 7;
            doc.setFont('helvetica', 'bold'); doc.setFontSize(valor.length <= 26 ? 15 : 11); doc.setTextColor(...COR.tinta);
            doc.text(valor, cx, yy, { align: 'center' }); yy += 5.5;
            doc.setFont('helvetica', 'normal'); doc.setFontSize(8); doc.setTextColor(...COR.tintaSec);
            subs.forEach(s => { doc.text(doc.splitTextToSize(String(s), w - 10)[0], cx, yy, { align: 'center' }); yy += 4.2; });
            return;
        }

        const px = x + 5;
        doc.setFont('helvetica', 'bold'); doc.setFontSize(7.5); doc.setTextColor(...COR.muted);
        doc.text(String(titulo).toUpperCase(), px, y + 6.5);
        const grande = valor.length <= 13;
        doc.setFont('helvetica', 'bold'); doc.setFontSize(grande ? 20 : 14); doc.setTextColor(...COR.tinta);
        doc.text(valor, px, y + (grande ? 16.5 : 15));
        if (subs.length) {
            doc.setFont('helvetica', 'normal'); doc.setFontSize(8); doc.setTextColor(...COR.tintaSec);
            let yy = y + (grande ? 22 : 20);
            subs.forEach(s => { doc.text(doc.splitTextToSize(String(s), w - 8)[0], px, yy); yy += 4; });
        }
    }

    // Gráfico de barras horizontais. itens: [{label, valor, cor?}] na ordem de exibição.
    // cor: cor padrão das barras (acento semântico do gráfico); item.cor sobrepõe se definida.
    function desenharBarras(doc, x, y, w, h, titulo, itens, fmt, cor) {
        fmt = fmt || (v => String(v));
        cor = cor || COR.azul;
        tituloSecao(doc, x, y + 4, w, titulo);
        const topo = y + 10;
        const areaH = h - 10;
        if (!itens.length) return;
        // Rótulo ocupa metade da largura (útil sobretudo no gráfico de largura total,
        // para caber o nome completo do tipo de documento em vez de truncar em "PETIÇÃO DE")
        const rotuloW = Math.min(w * 0.5, 120);
        const valorW = 13;
        const barX = x + rotuloW;
        const barMaxW = Math.max(6, w - rotuloW - valorW);
        const maxVal = Math.max(...itens.map(i => i.valor)) || 1;
        const linhaH = Math.min(9, areaH / itens.length);
        const barH = Math.max(3, linhaH * 0.58);

        doc.setFontSize(8);
        itens.forEach((it, i) => {
            const meio = topo + i * linhaH + linhaH / 2;
            doc.setFont('helvetica', 'normal'); doc.setTextColor(...COR.tintaSec);
            doc.text(doc.splitTextToSize(it.label, rotuloW - 3)[0], x, meio + 1.2);
            const bw = Math.max(0.6, (it.valor / maxVal) * barMaxW);
            doc.setFillColor(...(it.cor || cor));
            doc.roundedRect(barX, meio - barH / 2, bw, barH, 0.6, 0.6, 'F');
            doc.setFont('helvetica', 'bold'); doc.setTextColor(...COR.tinta);
            doc.text(fmt(it.valor), barX + bw + 2, meio + 1.2);
        });
    }

    // Cores de severidade por faixa (até 30d / 31-90d / mais de 90d), usadas como
    // marcador de ponto ao lado do rótulo — reforça a leitura sem depender só das barras.
    const COR_SEVERIDADE = [COR.aqua, COR.ambar, COR.vermelho];

    // Barras agrupadas por faixa: duas sub-barras (prioritários x normais) por linha, com
    // legenda e uma legenda de "como ler" — o vermelho é sempre prioritário, nunca a faixa.
    function desenharBarrasFaixas(doc, x, y, w, h, titulo, faixas) {
        tituloSecao(doc, x, y + 4, w, titulo, COR.ambar);
        const legY = y + 10;
        doc.setFillColor(...COR_PRIORITARIO); doc.rect(x, legY - 2.4, 3, 3, 'F');
        doc.setFont('helvetica', 'normal'); doc.setFontSize(7); doc.setTextColor(...COR.tintaSec);
        doc.text('Prioritários', x + 4, legY);
        const q2 = x + 28;
        doc.setFillColor(...COR.azul); doc.rect(q2, legY - 2.4, 3, 3, 'F');
        doc.text('Normais', q2 + 4, legY);

        const topo = y + 15;
        const areaH = h - 15;
        const rotuloW = Math.min(36, w * 0.32);
        const valorW = 10;
        const barX = x + rotuloW;
        const barMaxW = Math.max(6, w - rotuloW - valorW);
        const maxVal = Math.max(1, ...faixas.map(f => Math.max(f.prioritarios, f.normais)));
        const linhaH = areaH / faixas.length;
        const subH = Math.max(2.4, linhaH * 0.24);

        doc.setFontSize(7.5);
        faixas.forEach((f, i) => {
            const base = topo + i * linhaH;
            const meio = base + linhaH / 2;
            doc.setFillColor(...(COR_SEVERIDADE[i] || COR.muted));
            doc.circle(x + 1.4, meio - 0.4, 1.1, 'F');
            doc.setFont('helvetica', 'normal'); doc.setTextColor(...COR.tintaSec);
            doc.text(doc.splitTextToSize(f.label, rotuloW - 6)[0], x + 4, meio + 1);
            // prioritários (acima)
            const wp = Math.max(0.5, (f.prioritarios / maxVal) * barMaxW);
            doc.setFillColor(...COR_PRIORITARIO);
            doc.roundedRect(barX, meio - subH - 1, wp, subH, 0.5, 0.5, 'F');
            doc.setFont('helvetica', 'bold'); doc.setTextColor(...COR.tinta);
            doc.text(String(f.prioritarios), barX + wp + 1.5, meio - 1.4);
            // normais (abaixo)
            const wn = Math.max(0.5, (f.normais / maxVal) * barMaxW);
            doc.setFillColor(...COR.azul);
            doc.roundedRect(barX, meio + 1, wn, subH, 0.5, 0.5, 'F');
            doc.text(String(f.normais), barX + wn + 1.5, meio + subH + 0.6);
        });

        doc.setFont('helvetica', 'italic'); doc.setFontSize(6.6); doc.setTextColor(...COR.muted);
        doc.text('Como ler: cada faixa separa processos prioritários (vermelho) dos normais (azul), pelo tempo de espera.', x, y + h - 0.5);
    }

    // comIndice: se true, desenha um link "Voltar ao Índice" centralizado no rodapé,
    // apontando para a página 1 (usado apenas no PDF conjunto). Evita o glifo "↑" — fora
    // da codificação padrão das fontes do jsPDF e corrompe o texto renderizado.
    function desenharRodape(doc, titulo, quando, pw, ph, m, comIndice) {
        doc.setFont('helvetica', 'normal'); doc.setFontSize(8); doc.setTextColor(...COR.muted);
        doc.text(`${titulo}  •  Página ${doc.internal.getNumberOfPages()}`, m, ph - 6);
        doc.text(quando, pw - m, ph - 6, { align: 'right' });
        if (comIndice) {
            doc.setFont('helvetica', 'bold'); doc.setTextColor(...COR.azul);
            doc.textWithLink('Voltar ao Índice', pw / 2 - 16, ph - 6, { pageNumber: 1 });
        }
    }

    // Monta as páginas de RESUMO (geral + por competência) de um relatório genérico
    // (Retorno/Juntadas) dentro de um documento jsPDF já criado. ehPrimeiraSecao=false
    // começa em página nova (uso no conjunto). comIndice ativa o link de rodapé.
    function montarResumoGenerico(doc, dados, cfg, ehPrimeiraSecao, comIndice) {
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
        const gap = 6;

        // Desenha uma página de resumo (KPIs + gráficos) para um subconjunto de dados.
        function desenharPaginaResumo(sub, contexto, primeira) {
            if (!primeira) doc.addPage();

            // Cabeçalho: título, competência em destaque (com quebra), e a data em linha própria
            let hy = m + 2;
            doc.setFont('helvetica', 'bold'); doc.setFontSize(16); doc.setTextColor(...COR.tinta);
            doc.text(p.titulo, m, hy);
            hy += 8;
            if (contexto.competencia) {
                doc.setFont('helvetica', 'bold'); doc.setFontSize(11.5); doc.setTextColor(...COR.azul);
                const linhas = doc.splitTextToSize('Competência: ' + contexto.competencia, uw);
                doc.text(linhas, m, hy);
                hy += linhas.length * 5.2 + 1.5;
            } else if (contexto.rotulo) {
                doc.setFont('helvetica', 'bold'); doc.setFontSize(11.5); doc.setTextColor(...COR.azul);
                doc.text(contexto.rotulo, m, hy);
                hy += 7;
            }
            doc.setFont('helvetica', 'normal'); doc.setFontSize(9); doc.setTextColor(...COR.tintaSec);
            doc.text(`Extraído em ${hoje} às ${hora}  •  ${sub.length} registro(s)`, m, hy);
            hy += 3;
            doc.setDrawColor(...COR.azul); doc.setLineWidth(0.5); doc.line(m, hy, pw - m, hy);

            // KPIs numéricos (média por dia só quando o relatório define mediaLabel)
            const kY = hy + 6;
            const prio = contarPrioritarios(sub);
            const kpis = [
                { titulo: p.atosTitulo, valor: String(sub.length), subs: [], acento: COR.azul },
                { titulo: 'Prioritários pendentes', valor: String(prio), subs: [`${sub.length ? Math.round(prio / sub.length * 100) : 0}% do total`], acento: COR.vermelho },
            ];
            if (p.mediaLabel) {
                const media = mediaPorDia(sub, p.dataCampo);
                kpis.push({ titulo: 'Média por dia', valor: media ? media.toLocaleString('pt-BR', { maximumFractionDigits: 1 }) : '—', subs: [p.mediaLabel], acento: COR.aqua });
            }
            const kW = (uw - (kpis.length - 1) * gap) / kpis.length;
            kpis.forEach((k, i) => desenharCard(doc, m + i * (kW + gap), kY, kW, 28, k.titulo, k.valor, k.subs, false, k.acento));

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
            desenharCard(doc, m, aY, uw, 28, p.dataTitulo, valAntigo, subsAntigo, true, COR.ambar);

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
                else desenharBarras(doc, cx, cy, cw, chartH, c.titulo, c.itens, undefined, COR.aqua);
            });

            desenharRodape(doc, p.titulo, carimbo, pw, ph, m, comIndice);
        }

        // ═══ RESUMO GERAL ═══
        desenharPaginaResumo(dados, { rotulo: 'Resumo geral' }, ehPrimeiraSecao);

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
    }

    // Monta a TABELA DISCRIMINADA de um relatório genérico dentro de um documento jsPDF
    // já criado (sempre inicia em página nova). Retorna o número da página inicial da
    // tabela (para o índice/bookmarks do PDF conjunto).
    function montarTabelaGenerico(doc, dados, cfg, comIndice) {
        const p = cfg.pdf;
        const now = Date.now();
        const agora = new Date();
        const pw = doc.internal.pageSize.getWidth();
        const ph = doc.internal.pageSize.getHeight();
        const m = 12;
        const hoje = agora.toLocaleDateString('pt-BR');
        const hora = agora.toLocaleTimeString('pt-BR', { hour: '2-digit', minute: '2-digit' });
        const carimbo = `${hoje} ${hora}`;

        const ordenados = dados.slice().sort((a, b) => {
            const ta = parseDataBR(a[p.dataCampo]); const tb = parseDataBR(b[p.dataCampo]);
            return (ta == null ? Infinity : ta) - (tb == null ? Infinity : tb);
        });

        doc.addPage();
        const paginaInicial = doc.internal.getNumberOfPages();
        tituloSecao(doc, m, m + 3, pw - 2 * m, p.tabelaTitulo || 'Tabela discriminada');
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
            styles: { font: 'helvetica', fontSize: 7.5, cellPadding: 1.6, textColor: COR.tintaSec,
                      lineColor: COR.grade, lineWidth: 0.1, overflow: 'linebreak', valign: 'middle' },
            headStyles: { fillColor: COR.azul, textColor: [255, 255, 255], fontStyle: 'bold', fontSize: 8 },
            alternateRowStyles: { fillColor: COR.cartao },
            columnStyles,
            // Realça o número do processo dos prioritários
            didParseCell: (data) => {
                if (data.section === 'body' && data.column.index === idxProcesso && ordenados[data.row.index] && ordenados[data.row.index].prioritario) {
                    data.cell.styles.textColor = COR_PRIORITARIO;
                    data.cell.styles.fontStyle = 'bold';
                }
            },
            didDrawPage: () => desenharRodape(doc, p.titulo, carimbo, pw, ph, m, comIndice),
        });

        return paginaInicial;
    }

    function novoDocPDF() {
        const ctor = (window.jspdf && window.jspdf.jsPDF) || window.jsPDF;
        if (!ctor) throw new Error('biblioteca jsPDF não carregada');
        const doc = new ctor({ orientation: 'portrait', unit: 'mm', format: 'a4' });
        if (typeof doc.autoTable !== 'function') throw new Error('plugin autoTable não carregado');
        return doc;
    }

    function gerarPDF(dados, cfg) {
        const doc = novoDocPDF();
        montarResumoGenerico(doc, dados, cfg, true, false);
        const pgTabela = montarTabelaGenerico(doc, dados, cfg, false);
        doc.outline.add(null, 'Resumo', { pageNumber: 1 });
        doc.outline.add(null, 'Tabela detalhada', { pageNumber: pgTabela });
        baixarBlob(doc.output('blob'), `${cfg.nomeArquivo}_${dataArquivo()}.pdf`);
    }

    // Resolve como montar o resumo/tabela de uma seção do PDF conjunto, dado o cfg do
    // relatório (genérico via cfg.pdf, ou o caso especial do Tempo Médio).
    function descreverSecaoPDF(cfg) {
        if (cfg === CFG_TEMPOMEDIO) {
            return {
                rotulo: 'Tempo Médio de Cumprimento',
                montarResumo: (doc, dados, primeira, comIndice) => montarResumoTempoMedio(doc, dados, primeira, comIndice),
                montarTabela: (doc, dados, comIndice) => montarTabelaTempoMedio(doc, dados, comIndice),
            };
        }
        return {
            rotulo: cfg.pdf.titulo,
            montarResumo: (doc, dados, primeira, comIndice) => montarResumoGenerico(doc, dados, cfg, primeira, comIndice),
            montarTabela: (doc, dados, comIndice) => montarTabelaGenerico(doc, dados, cfg, comIndice),
        };
    }

    // Capa + índice clicável (página 1). Retorna as linhas desenhadas — cada uma com a
    // posição Y, a seção e o tipo ("Resumo"/"Tabela detalhada") — para depois (após montar
    // todo o conteúdo e já sabendo os números de página) voltar e completar os links.
    function desenharCapaIndice(doc, secoes, agora) {
        const pw = doc.internal.pageSize.getWidth();
        const ph = doc.internal.pageSize.getHeight();
        const m = 16;
        const hoje = agora.toLocaleDateString('pt-BR');
        const hora = agora.toLocaleTimeString('pt-BR', { hour: '2-digit', minute: '2-digit' });

        doc.setFillColor(...COR.azul); doc.rect(0, 0, pw, 26, 'F');
        doc.setFont('helvetica', 'bold'); doc.setFontSize(20); doc.setTextColor(255, 255, 255);
        doc.text('Relatório Conjunto', m, 17);
        doc.setFont('helvetica', 'normal'); doc.setFontSize(10); doc.setTextColor(...COR.tintaSec);
        doc.text(`Projudi — TJPR  •  Extraído em ${hoje} às ${hora}`, m, 36);

        let y = 48;
        doc.setFont('helvetica', 'bold'); doc.setFontSize(11); doc.setTextColor(...COR.tinta);
        doc.text('Conteúdo', m, y); y += 3;
        doc.setDrawColor(...COR.azul); doc.setLineWidth(0.5); doc.line(m, y, m + 26, y); y += 8;

        const linhas = [];
        doc.setFontSize(10);
        secoes.forEach((s, i) => {
            doc.setFont('helvetica', 'bold'); doc.setFontSize(11); doc.setTextColor(...COR.tinta);
            doc.text(`${i + 1}. ${s.rotulo}`, m, y);
            doc.setFont('helvetica', 'normal'); doc.setFontSize(9); doc.setTextColor(...COR.muted);
            doc.text(`${s.dados.length} registro(s)`, pw - m, y, { align: 'right' });
            y += 6;
            ['Resumo', 'Tabela detalhada'].forEach(tipo => {
                doc.setFont('helvetica', 'normal'); doc.setFontSize(9.5); doc.setTextColor(...COR.azul);
                doc.text(`•  ${tipo}`, m + 4, y);
                linhas.push({ y, secaoIdx: i, tipo });
                y += 5.5;
            });
            y += 3;
        });

        doc.setFont('helvetica', 'italic'); doc.setFontSize(8); doc.setTextColor(...COR.muted);
        doc.text('Todos os resumos vêm primeiro; as tabelas discriminadas ficam ao final. Clique num item para navegar.', m, ph - 16);
        return linhas;
    }

    // PDF único com as seções na ordem informada. secoes: [{ dados, cfg }, ...]. Estrutura:
    // capa+índice clicável (pág. 1) → todos os resumos → todas as tabelas discriminadas,
    // com marcadores (bookmarks) no leitor e link "Voltar ao Índice" no rodapé das tabelas.
    function gerarPDFConjunto(secoesEntrada) {
        const doc = novoDocPDF();
        const agora = new Date();
        const secoes = secoesEntrada.map(s => Object.assign({ dados: s.dados }, descreverSecaoPDF(s.cfg)));

        // PASSO A: capa + índice (números de página ainda em branco)
        const linhasIndice = desenharCapaIndice(doc, secoes, agora);

        // PASSO B: bloco de RESUMOS, depois bloco de TABELAS — todos os resumos antes de
        // qualquer tabela, conforme pedido. Marcadores agrupam por bloco.
        const paginaResumo = [], paginaTabela = [];
        const bmResumos = doc.outline.add(null, 'Resumos', { pageNumber: 2 });
        secoes.forEach(s => {
            const pg = doc.internal.getNumberOfPages() + 1;
            s.montarResumo(doc, s.dados, false, true);
            paginaResumo.push(pg);
            doc.outline.add(bmResumos, s.rotulo, { pageNumber: pg });
        });
        const primeiraTabelaPg = doc.internal.getNumberOfPages() + 1;
        const bmTabelas = doc.outline.add(null, 'Tabelas detalhadas', { pageNumber: primeiraTabelaPg });
        secoes.forEach(s => {
            const pg = s.montarTabela(doc, s.dados, true);
            paginaTabela.push(pg);
            doc.outline.add(bmTabelas, s.rotulo, { pageNumber: pg });
        });

        // PASSO C: volta à página 1 e completa os números + links do índice
        doc.setPage(1);
        const pw = doc.internal.pageSize.getWidth();
        const m = 16;
        linhasIndice.forEach(l => {
            const pg = l.tipo === 'Resumo' ? paginaResumo[l.secaoIdx] : paginaTabela[l.secaoIdx];
            doc.setDrawColor(...COR.grade); doc.setLineWidth(0.2); doc.setLineDashPattern([0.5, 1], 0);
            doc.line(m + 40, l.y - 0.8, pw - m - 10, l.y - 0.8);
            doc.setLineDashPattern([], 0);
            doc.setFont('helvetica', 'normal'); doc.setFontSize(9.5); doc.setTextColor(...COR.azul);
            doc.textWithLink(`•  ${l.tipo}`, m + 4, l.y, { pageNumber: pg });
            doc.setTextColor(...COR.tintaSec); doc.text(String(pg), pw - m, l.y, { align: 'right' });
        });

        baixarBlob(doc.output('blob'), `relatorio_conjunto_projudi_${dataArquivo()}.pdf`);
    }

    // ── PDF do relatório de Tempo Médio de Cumprimento ──────────────────────────

    // Média de d[valorCampo] agrupada por d[campo]; top-N por média desc, resto vira "Outros"
    // (com a média ponderada pelo número de registros de cada grupo restante).
    function agregarMedia(dados, campo, valorCampo, topN) {
        const mapa = new Map();
        dados.forEach(d => {
            const v = d[valorCampo];
            if (v == null) return;
            const k = (d[campo] || '').trim() || '(vazio)';
            if (!mapa.has(k)) mapa.set(k, { soma: 0, n: 0 });
            const o = mapa.get(k); o.soma += v; o.n += 1;
        });
        let arr = [...mapa.entries()].map(([label, o]) => ({ label, valor: o.soma / o.n, n: o.n }))
            .sort((a, b) => b.valor - a.valor);
        if (arr.length > topN) {
            const resto = arr.slice(topN);
            const nResto = resto.reduce((s, i) => s + i.n, 0);
            const somaResto = resto.reduce((s, i) => s + i.valor * i.n, 0);
            arr = arr.slice(0, topN);
            if (nResto) arr.push({ label: 'Outros', valor: somaResto / nResto, n: nResto });
        }
        return arr;
    }

    function mediaSimples(itens, campo) {
        const validos = itens.filter(d => d[campo] != null);
        if (!validos.length) return null;
        return validos.reduce((s, d) => s + d[campo], 0) / validos.length;
    }

    function fmtDias(v) {
        return v == null ? '—' : `${v.toFixed(1)} dia${Math.abs(v - 1) < 0.05 ? '' : 's'}`;
    }

    // Lista dos processos mais demorados: linha dupla (processo em negrito + classe em
    // cinza) à esquerda e uma barra proporcional aos dias à direita.
    function desenharTopDemorados(doc, x, y, w, h, titulo, itens) {
        tituloSecao(doc, x, y + 4, w, titulo, COR.vermelho);
        const topo = y + 10;
        const areaH = h - 10;
        if (!itens.length) return;
        const labelW = Math.min(w * 0.62, 120);
        const barX = x + labelW;
        const barMaxW = Math.max(6, w - labelW - 14);
        const maxVal = Math.max(...itens.map(i => i.dias)) || 1;
        const linhaH = areaH / itens.length;
        const barH = Math.max(2.2, linhaH * 0.3);

        itens.forEach((it, i) => {
            const rowY = topo + i * linhaH;
            const cor = it.prioritario ? COR_PRIORITARIO : COR.tinta;
            doc.setFont('helvetica', 'bold'); doc.setFontSize(7.5); doc.setTextColor(...cor);
            doc.text(doc.splitTextToSize(it.processo || '', labelW - 3)[0], x, rowY + 3.3);
            doc.setFont('helvetica', 'normal'); doc.setFontSize(6.8); doc.setTextColor(...COR.muted);
            doc.text(doc.splitTextToSize(it.classe || '', labelW - 3)[0], x, rowY + 6.8);

            const meio = rowY + linhaH / 2;
            const bw = Math.max(0.6, (it.dias / maxVal) * barMaxW);
            doc.setFillColor(...(it.prioritario ? COR_PRIORITARIO : COR.azul));
            doc.roundedRect(barX, meio - barH / 2, bw, barH, 0.5, 0.5, 'F');
            doc.setFont('helvetica', 'bold'); doc.setFontSize(7.5); doc.setTextColor(...COR.tinta);
            doc.text(`${it.dias} dia${it.dias === 1 ? '' : 's'}`, barX + bw + 2, meio + 1.2);
        });
    }

    const TITULO_TEMPOMEDIO = 'Tempo Médio de Cumprimento';

    function gerarPDFTempoMedio(dados) {
        const doc = novoDocPDF();
        montarResumoTempoMedio(doc, dados, true, false);
        const pgTabela = montarTabelaTempoMedio(doc, dados, false);
        doc.outline.add(null, 'Resumo', { pageNumber: 1 });
        doc.outline.add(null, 'Tabela detalhada', { pageNumber: pgTabela });
        baixarBlob(doc.output('blob'), `tempo_medio_projudi_${dataArquivo()}.pdf`);
    }

    // Página de RESUMO (KPIs + gráficos) do relatório de Tempo Médio, dentro de um doc
    // jsPDF já existente. ehPrimeiraSecao=false começa em página nova (uso no conjunto).
    function montarResumoTempoMedio(doc, dados, ehPrimeiraSecao, comIndice) {
        if (!ehPrimeiraSecao) doc.addPage();
        const agora = new Date();
        const pw = doc.internal.pageSize.getWidth();
        const ph = doc.internal.pageSize.getHeight();
        const m = 12;
        const uw = pw - 2 * m;
        const hoje = agora.toLocaleDateString('pt-BR');
        const hora = agora.toLocaleTimeString('pt-BR', { hour: '2-digit', minute: '2-digit' });

        const validos = dados.filter(d => d.dias != null);
        const prioritarios = validos.filter(d => d.prioritario);
        const naoPrioritarios = validos.filter(d => !d.prioritario);
        const geral = mediaSimples(validos, 'dias');
        const mediaPrio = mediaSimples(prioritarios, 'dias');
        const mediaNaoPrio = mediaSimples(naoPrioritarios, 'dias');
        const maisDemorado = validos.slice().sort((a, b) => b.dias - a.dias)[0] || null;

        // Período de referência (salvo quando o usuário preencheu o formulário)
        let periodoStr = '';
        try {
            const per = JSON.parse(store.getItem('projudi_tempomedio_periodo') || '{}');
            if (per.ini || per.fim) periodoStr = `${per.ini || '?'} a ${per.fim || '?'}`;
        } catch (e) {}

        // Decisões ainda não cumpridas (dtCartorio vazia = cartório não analisou ainda)
        const naoCumpridas = dados.filter(d => !d.dtCartorio);
        const maisAntigaNC = naoCumpridas.reduce((best, d) => {
            const ts = parseDataBR(d.dtAnalise);
            return ts != null && (best === null || ts < best.ts) ? { ts, str: d.dtAnalise } : best;
        }, null);

        doc.setFont('helvetica', 'bold'); doc.setFontSize(16); doc.setTextColor(...COR.tinta);
        doc.text(TITULO_TEMPOMEDIO, m, m + 2);
        doc.setFont('helvetica', 'normal'); doc.setFontSize(9); doc.setTextColor(...COR.tintaSec);
        let subtitulo = `Extraído em ${hoje} às ${hora}  •  ${dados.length} registro(s) analisado(s)`;
        if (periodoStr) subtitulo += `  •  Período: ${periodoStr}`;
        doc.text(subtitulo, m, m + 8);
        doc.setDrawColor(...COR.azul); doc.setLineWidth(0.5); doc.line(m, m + 11, pw - m, m + 11);

        const gap = 6;
        // Linha 1: 4 KPIs centralizados — analisados / tempo médio geral / prioritários / não cumpridas
        const kY = m + 16;
        const kW4 = (uw - 3 * gap) / 4;
        const prioPct = validos.length ? Math.round(prioritarios.length / validos.length * 100) : 0;
        desenharCard(doc, m,               kY, kW4, 28, 'Registros analisados', String(dados.length), [], true, COR.azul);
        desenharCard(doc, m + kW4 + gap,   kY, kW4, 28, 'Tempo médio geral', fmtDias(geral), [], true, COR.aqua);
        desenharCard(doc, m + 2*(kW4+gap), kY, kW4, 28, 'Prioritários', String(prioritarios.length), [`${prioPct}% do total`], true, COR.vermelho);
        desenharCard(doc, m + 3*(kW4+gap), kY, kW4, 28, 'Não cumpridas', String(naoCumpridas.length),
            [maisAntigaNC ? `mais antiga: ${maisAntigaNC.str}` : ''], true, COR.ambar);

        // Linha 2: tempo médio prioritários vs não prioritários (centralizados)
        const k2Y = kY + 28 + gap;
        const kW2 = (uw - gap) / 2;
        desenharCard(doc, m,           k2Y, kW2, 26, 'Tempo médio — Prioritários',     fmtDias(mediaPrio),    [`${prioritarios.length} processo(s)`], true, COR.vermelho);
        desenharCard(doc, m + kW2+gap, k2Y, kW2, 26, 'Tempo médio — Não prioritários', fmtDias(mediaNaoPrio), [`${naoPrioritarios.length} processo(s)`], true, COR.azul);

        // Card largo: processo com cumprimento mais demorado (centralizado)
        const k3Y = k2Y + 26 + gap; // k2Y usa h=26
        let valMD = '—', subsMD = ['Nenhum registro com datas válidas'];
        if (maisDemorado) {
            valMD = `${maisDemorado.dias} dia${maisDemorado.dias === 1 ? '' : 's'}`;
            subsMD = [
                `Processo ${maisDemorado.processo}${maisDemorado.prioritario ? '  — PRIORITÁRIO' : ''}`,
                maisDemorado.classe || '',
            ];
        }
        desenharCard(doc, m, k3Y, uw, 26, 'Processo com cumprimento mais demorado', valMD, subsMD, true, COR.vermelho);

        // Gráficos 1 e 2 (largura total, empilhados): dividem o espaço vertical restante da
        // página de forma que a soma das duas alturas + o espaçamento sempre caiba, sem
        // depender de contagens fixas (evita transbordar para a área do rodapé).
        const chartY = k3Y + 26 + gap + 2;
        const chartGap = 8;
        const disponivel = ph - m - chartY - 14; // reserva a faixa do rodapé
        const top10 = validos.slice().sort((a, b) => b.dias - a.dias).slice(0, 10)
            .map(d => ({ processo: d.processo, classe: d.classe, dias: d.dias, prioritario: d.prioritario }));
        const chart1Desejado = Math.max(30, top10.length * 8 + 8);
        const chart1H = Math.min(chart1Desejado, Math.max(30, disponivel - chartGap - 40));
        const chart2H = Math.max(30, disponivel - chart1H - chartGap);
        desenharTopDemorados(doc, m, chartY, uw, chart1H, 'Processos com cumprimento mais demorado', top10);

        const chart2Y = chartY + chart1H + chartGap;
        const porClasse = agregarMedia(validos, 'classe', 'dias', 12);
        desenharBarras(doc, m, chart2Y, uw, chart2H, 'Tempo médio por Classe Processual', porClasse, fmtDias, COR.aqua);

        desenharRodape(doc, TITULO_TEMPOMEDIO, `${hoje} ${hora}`, pw, ph, m, comIndice);
    }

    // Tabela discriminada do relatório de Tempo Médio (sempre inicia em página nova).
    // Retorna o número da página inicial (para o índice/bookmarks do PDF conjunto).
    function montarTabelaTempoMedio(doc, dados, comIndice) {
        const agora = new Date();
        const pw = doc.internal.pageSize.getWidth();
        const ph = doc.internal.pageSize.getHeight();
        const m = 12;
        const hoje = agora.toLocaleDateString('pt-BR');
        const hora = agora.toLocaleTimeString('pt-BR', { hour: '2-digit', minute: '2-digit' });

        // Do maior número de dias para o menor (sem dados vai ao final)
        const ordenados = dados.slice().sort((a, b) => {
            const da = a.dias == null ? -Infinity : a.dias;
            const db = b.dias == null ? -Infinity : b.dias;
            return db - da;
        });

        doc.addPage();
        const paginaInicial = doc.internal.getNumberOfPages();
        tituloSecao(doc, m, m + 3, pw - 2 * m, 'Tabela discriminada — todos os resultados');

        const colunas = [
            { header: 'Processo', width: 30, get: (d) => d.processo },
            { header: 'Dt. Análise', width: 22, get: (d) => d.dtAnalise },
            { header: 'Dt. Análise Cartório', width: 26, get: (d) => d.dtCartorio },
            { header: 'Tipo de Conclusão', width: 30, get: (d) => d.tipoConclusao },
            { header: 'Classe Processual', width: 42, get: (d) => d.classe },
            { header: 'Dias p/ Cumprimento', width: 20, get: (d) => (d.dias == null ? '' : String(d.dias)) },
        ];
        const columnStyles = {};
        colunas.forEach((c, i) => { columnStyles['k' + i] = { cellWidth: c.width }; });
        const idxProcesso = 0;

        doc.autoTable({
            columns: colunas.map((c, i) => ({ header: c.header, dataKey: 'k' + i })),
            body: ordenados.map(d => {
                const o = {};
                colunas.forEach((c, i) => { o['k' + i] = String(c.get(d) ?? ''); });
                return o;
            }),
            startY: m + 8,
            margin: { left: m, right: m, top: m, bottom: 14 },
            theme: 'grid',
            styles: { font: 'helvetica', fontSize: 7.5, cellPadding: 1.6, textColor: COR.tintaSec,
                      lineColor: COR.grade, lineWidth: 0.1, overflow: 'linebreak', valign: 'middle' },
            headStyles: { fillColor: COR.azul, textColor: [255, 255, 255], fontStyle: 'bold', fontSize: 8 },
            alternateRowStyles: { fillColor: COR.cartao },
            columnStyles,
            didParseCell: (data) => {
                if (data.section === 'body' && data.column.index === idxProcesso && ordenados[data.row.index] && ordenados[data.row.index].prioritario) {
                    data.cell.styles.textColor = COR_PRIORITARIO;
                    data.cell.styles.fontStyle = 'bold';
                }
            },
            didDrawPage: () => desenharRodape(doc, TITULO_TEMPOMEDIO, `${hoje} ${hora}`, pw, ph, m, comIndice),
        });

        return paginaInicial;
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
        #btn-preencher-pesquisar-tm { background-color: #1e6b1e; border-color: #145214; }
        .projudi-btn:disabled { background-color: #999; border-color: #777; cursor: not-allowed; }
        .projudi-select {
            margin-left: 6px; padding: 2px 4px; font-size: 0.85em;
            font-family: Verdana, Arial, sans-serif; border-radius: 3px; border: 1px solid #999;
        }
        #exportar-status { font-size: 0.8em; color: #555; margin-left: 8px; font-family: Verdana, Arial, sans-serif; }
    `);

    // Detecta qual relatório está na tela pelo cabeçalho da resultTable; se não houver
    // tabela reconhecível, tenta pela URL (útil na tela de busca da página de juntadas).
    function detectarConfig() {
        const thead = document.querySelector('table.resultTable thead');
        const cab = thead ? thead.textContent : '';
        let cfg = null;
        if (CFG_TEMPOMEDIO.detecta(cab)) cfg = CFG_TEMPOMEDIO;
        else if (CFG_JUNTADAS.detecta(cab)) cfg = CFG_JUNTADAS;
        else if (CFG_RETORNO.detecta(cab)) cfg = CFG_RETORNO;
        else if (CFG_CONCLUSOES.detecta(cab)) cfg = CFG_CONCLUSOES;
        else if (/analisarJuntada\.do/i.test(location.pathname + location.search)) cfg = CFG_JUNTADAS;
        console.log(`[Projudi] detectarConfig — thead=${!!thead} cfg=${cfg ? cfg.prefixo : 'null'} cab="${cab.slice(0,80).replace(/\s+/g,' ')}"`);
        return cfg;
    }

    // Tela de filtros do relatório "Estatísticas de Conclusões" (tempo médio), antes da pesquisa.
    function formularioTempoMedio() {
        const form = document.getElementById('estatisticaConclusaoForm');
        return form && form.querySelector('input[name="situacao"]') ? form : null;
    }

    // Períodos pré-configurados para a data inicial do relatório de Tempo Médio.
    const PERIODOS_TEMPOMEDIO = [
        { id: '1m',  rotulo: 'Último mês inteiro', calc: () => { const d = new Date(); d.setMonth(d.getMonth() - 1); return d; } },
        { id: '1a',  rotulo: 'Último 1 ano',        calc: () => { const d = new Date(); d.setFullYear(d.getFullYear() - 1); return d; } },
        { id: '2a',  rotulo: 'Últimos 2 anos',       calc: () => { const d = new Date(); d.setFullYear(d.getFullYear() - 2); return d; } },
        { id: '3a',  rotulo: 'Últimos 3 anos',       calc: () => { const d = new Date(); d.setFullYear(d.getFullYear() - 3); return d; } },
    ];

    function formatarDataBR(d) {
        const dd = String(d.getDate()).padStart(2, '0');
        const mm = String(d.getMonth() + 1).padStart(2, '0');
        return `${dd}/${mm}/${d.getFullYear()}`;
    }

    // Marca Situação=Analisadas, Tipo=Analítico, define a data inicial conforme o período
    // escolhido (select ao lado do botão) e clica em Pesquisar. O diagnóstico anterior (3s)
    // apontava falha, mas era falso negativo: o site do Projudi é lento e a navegação real
    // ainda estava em andamento quando o diagnóstico rodou — por isso o timeout de checagem
    // agora é bem mais longo.
    function preencherEPesquisarTempoMedio(periodoId) {
        const form = formularioTempoMedio();
        if (!form) return;

        const periodo = PERIODOS_TEMPOMEDIO.find(p => p.id === periodoId) || PERIODOS_TEMPOMEDIO[3];

        const radioAnalisadas = form.querySelector('input[name="situacao"][value="A"]');
        const radioAnalitico = form.querySelector('input[name="analitico"][value="true"]');
        console.log(`[Projudi TM] radioAnalisadas encontrado=${!!radioAnalisadas} radioAnalitico encontrado=${!!radioAnalitico} período="${periodo.rotulo}"`);
        if (radioAnalisadas) radioAnalisadas.checked = true;
        if (radioAnalitico) radioAnalitico.checked = true;

        const campoInicio = form.querySelector('input[name="dataInicio"]');
        if (campoInicio) {
            // Campos de data vêm com "disabled" no HTML original: se ficarem desabilitados,
            // o navegador NÃO os envia no submit. Precisamos habilitar antes de definir o valor.
            campoInicio.disabled = false;
            campoInicio.value = formatarDataBR(periodo.calc());
        }
        const campoFim = form.querySelector('input[name="dataFim"]');
        if (campoFim) campoFim.disabled = false; // também precisa ir habilitado para ser enviado
        console.log(`[Projudi TM] campos preenchidos — dataInicio="${campoInicio ? campoInicio.value : '?'}" dataFim="${campoFim ? campoFim.value : '?'}"`);

        // Salva o período para uso posterior no PDF
        store.setItem('projudi_tempomedio_periodo', JSON.stringify({
            ini: campoInicio ? campoInicio.value : '',
            fim: campoFim ? campoFim.value : '',
        }));

        // Sinaliza que, ao carregar a página de resultados, a extração deve iniciar
        // automaticamente (sem esta flag a página de resultados só renderiza os botões).
        store.setItem('projudi_tempomedio_auto_iniciar', '1');

        const btn = document.getElementById('searchButton') || form.querySelector('input[type="submit"]');
        console.log(`[Projudi TM] flag auto_iniciar definida; clicando em Pesquisar em 1,5s (btn encontrado=${!!btn})`);

        setTimeout(() => {
            console.log('[Projudi TM] clicando em Pesquisar — o site pode demorar para responder, aguarde.');
            if (btn && !btn.disabled) btn.click(); else form.submit();

            // Diagnóstico tardio (site é lento; não dispara nenhum reenvio, só informa).
            setTimeout(() => {
                const aindaNoFormulario = !!document.getElementById('estatisticaConclusaoForm');
                const temResultado = !!document.querySelector('table.resultTable');
                console.log(`[Projudi TM] diagnóstico 15s depois — aindaNoFormulario=${aindaNoFormulario} temResultado=${temResultado}`);
                if (aindaNoFormulario && !temResultado) {
                    console.warn('[Projudi TM] ainda sem resultado após 15s — o site pode estar lento; se persistir por muito mais tempo, clique em Pesquisar manualmente.');
                }
            }, 15000);
        }, 1500);
    }

    function injetarBotoes() {
        const buttonBar = document.querySelector('table.buttonBar td.buttons');
        if (!buttonBar) return;

        // Tela de filtros do relatório de Tempo Médio (ainda sem resultados): select do
        // período + botão de preencher+pesquisar; os botões de coleta/exportação fazem
        // sentido depois da busca.
        if (formularioTempoMedio() && !document.querySelector('table.resultTable')) {
            // Automação: se chegamos aqui vindos do fluxo de automação, preenche e
            // pesquisa sozinho (sem esperar clique manual), usando o período escolhido no painel.
            if (store.getItem(AUTO_ESTADO) === 'preenchendo_tempomedio') {
                const periodoTM = store.getItem('projudi_auto_periodo_tm') || '3a';
                console.log(`[Projudi TM] automação: preenchendo e pesquisando com período="${periodoTM}"`);
                store.setItem(AUTO_ESTADO, 'coletando_tempomedio');
                preencherEPesquisarTempoMedio(periodoTM);
                return;
            }

            const sel = document.createElement('select');
            sel.id = 'sel-periodo-tm';
            sel.className = 'projudi-select';
            sel.title = 'Período usado como data inicial da pesquisa';
            PERIODOS_TEMPOMEDIO.forEach(p => {
                const opt = document.createElement('option');
                opt.value = p.id;
                opt.textContent = p.rotulo;
                sel.appendChild(opt);
            });
            sel.value = '3a'; // padrão: últimos 3 anos
            buttonBar.appendChild(sel);

            const b = document.createElement('button');
            b.id = 'btn-preencher-pesquisar-tm';
            b.type = 'button';
            b.className = 'projudi-btn';
            b.title = 'Marca Analisadas + Analítico, define a data inicial conforme o período escolhido e pesquisa (o site pode demorar a responder)';
            b.textContent = 'Preencher e Pesquisar';
            b.onclick = () => preencherEPesquisarTempoMedio(sel.value);
            buttonBar.appendChild(b);
            return;
        }

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
        if (cfg.pdf || cfg.pdfCustom) {
            buttonBar.appendChild(mk('btn-pdf', 'Gera um PDF com painel, gráficos e a tabela completa', () => coletor.pdf()));
        }
        buttonBar.appendChild(mk('btn-limpar', 'Apaga os dados acumulados deste relatório', () => coletor.limpar(), 'Limpar'));

        const status = document.createElement('span');
        status.id = 'exportar-status';
        buttonBar.appendChild(status);

        const estadoAuto = store.getItem(AUTO_ESTADO);
        const relAtual = relatorioPorCfg(cfg);
        const querColetarAuto = !!relAtual && estadoAuto === 'coletando_' + relAtual.key && cfg !== CFG_TEMPOMEDIO;
        const autoIniciarTM = cfg === CFG_TEMPOMEDIO && store.getItem('projudi_tempomedio_auto_iniciar') === '1';

        console.log(`[Projudi] injetarBotoes — cfg=${cfg.prefixo} rodando=${coletor.rodando()} obsoleta=${coletor.obsoleta()} querColetarAuto=${querColetarAuto} autoIniciarTM=${autoIniciarTM}`);

        if (coletor.rodando() && !coletor.obsoleta()) {
            console.log('[Projudi] retomando coleta após reload de paginação');
            coletor.continuar(); // retoma após o reload da paginação
        } else if (querColetarAuto) {
            console.log('[Projudi] automação: iniciando coleta ao chegar no relatório');
            coletor.iniciar();   // automação: inicia a coleta ao chegar no relatório
        } else if (autoIniciarTM) {
            store.removeItem('projudi_tempomedio_auto_iniciar');
            console.log('[Projudi TM] flag auto_iniciar detectada — iniciando extração automaticamente');
            coletor.iniciar();   // início automático após o usuário clicar em "Pesquisar"
        } else {
            console.log('[Projudi] nenhuma coleta em andamento — renderizando botões');
            coletor.limparFlags(); // descarta flag de execução presa, mantendo os dados
            coletor.render();
        }
    }

    // ══════════════════════════ AUTOMAÇÃO ══════════════════════════
    // Fluxo: (início) -> Juntadas [extrai tudo] -> início -> Retorno [extrai tudo]
    //        -> início -> habilita "PDF conjunto". O estado persiste no localStorage,
    //        então dá para rodar de novo em outra Atuação e ACUMULAR várias competências.
    const AUTO_ESTADO = 'projudi_auto_estado';

    function desembrulharArray(valor) {
        let v = valor, t = 0;
        while (typeof v === 'string' && t < 5) { try { v = JSON.parse(v); } catch (e) { break; } t++; }
        return Array.isArray(v) ? v : null;
    }

    // Lê os dados acumulados de um relatório a partir do prefixo de armazenamento.
    function lerDadosDe(prefixo) {
        const n = parseInt(store.getItem(prefixo + 'num_paginas') || '0', 10);
        let dados = [];
        for (let i = 0; i < n; i++) {
            const b = store.getItem(prefixo + 'pagina_' + i);
            if (!b) continue;
            const parte = desembrulharArray(b);
            if (parte) dados = dados.concat(parte);
        }
        return dados;
    }

    // Relatórios disponíveis para a automação, na ordem padrão de execução. 'precisaPreencher'
    // marca o Tempo Médio, cuja página de destino é um formulário de filtros (não os
    // resultados diretamente) — precisa ser preenchido e pesquisado antes de coletar.
    const REPORTS_AUTOMACAO = [
        { key: 'juntadas',   cfg: CFG_JUNTADAS,   navAlvo: 'juntadas',   rotulo: 'Juntadas',             precisaPreencher: false },
        { key: 'retorno',    cfg: CFG_RETORNO,    navAlvo: 'retorno',    rotulo: 'Retorno de Conclusos',  precisaPreencher: false },
        { key: 'tempomedio', cfg: CFG_TEMPOMEDIO, navAlvo: 'tempomedio', rotulo: 'Tempo Médio',           precisaPreencher: true },
    ];
    function relatorioPorChave(key) { return REPORTS_AUTOMACAO.find(r => r.key === key); }
    function relatorioPorCfg(cfg) { return REPORTS_AUTOMACAO.find(r => r.cfg === cfg); }

    function lerFilaAutomacao() {
        try { return JSON.parse(store.getItem('projudi_auto_fila') || '[]'); } catch (e) { return []; }
    }

    // Procura um link de menu (por URL e/ou texto) no documento atual e nos frames pai/topo.
    function acharLinkMenu(urlRe, textoRe) {
        const docs = [document];
        try { if (window.parent && window.parent !== window) docs.push(window.parent.document); } catch (e) {}
        try { if (window.top && window.top !== window) docs.push(window.top.document); } catch (e) {}
        for (const d of docs) {
            for (const a of d.querySelectorAll('a[href]')) {
                if (urlRe && !urlRe.test(a.href)) continue;
                if (textoRe && !textoRe.test((a.textContent || '').trim())) continue;
                return a;
            }
        }
        return null;
    }

    function navegarMenu(alvo) {
        let link = null;
        if (alvo === 'juntadas') link = acharLinkMenu(/analisarJuntada\.do/i, null);
        else if (alvo === 'retorno') link = acharLinkMenu(/conclusao\.do/i, /retorno de processos conclusos/i);
        else if (alvo === 'tempomedio') link = acharLinkMenu(/conclusao\/estatistica\.do/i, null);
        else if (alvo === 'inicio') link = acharLinkMenu(null, /^in[íi]cio$/i);
        if (!link) { console.warn('[Auto Projudi] link de menu não encontrado:', alvo); return false; }
        link.click();
        return true;
    }

    // Chamado ao concluir a coleta de um relatório (pelo coletor). Marca o próximo estado
    // (próximo relatório da fila, ou "ir_fim") e tenta avançar (o próprio frame do relatório
    // costuma ter o menu; senão o poll do painel assume).
    function avancarAutomacao(cfg) {
        const estado = store.getItem(AUTO_ESTADO);
        const rel = relatorioPorCfg(cfg);
        if (!rel || estado !== 'coletando_' + rel.key) return;
        const fila = lerFilaAutomacao();
        const idx = fila.indexOf(rel.key);
        const prox = idx >= 0 ? fila[idx + 1] : undefined;
        store.setItem(AUTO_ESTADO, prox ? ('ir_' + prox) : 'ir_fim');
        setTimeout(passoAutomacao, 900);
    }

    // Executa o passo de navegação do estado atual. Pode ser chamado por qualquer frame;
    // uma trava de tempo evita cliques duplicados quando mais de um frame o executa.
    function passoAutomacao() {
        const estado = store.getItem(AUTO_ESTADO);
        if (!estado || estado === 'concluido') return;
        // Durante a coleta ou o preenchimento do formulário, quem conduz é a própria
        // página atual (injetarBotoes / preencherEPesquisarTempoMedio).
        if (estado.startsWith('coletando_') || estado === 'preenchendo_tempomedio') return;

        const agora = Date.now();
        const lock = parseInt(store.getItem('projudi_auto_lock') || '0', 10);
        if (agora - lock < 4000) return; // já houve tentativa recente em algum frame

        if (estado === 'ir_fim') {
            store.setItem('projudi_auto_lock', String(agora));
            store.setItem(AUTO_ESTADO, 'concluido');
            navegarMenu('inicio');
            return;
        }
        if (estado.startsWith('ir_')) {
            const key = estado.slice(3);
            const rel = relatorioPorChave(key);
            if (!rel) { console.warn('[Auto Projudi] relatório desconhecido no estado', estado); return; }
            store.setItem('projudi_auto_lock', String(agora));
            if (navegarMenu(rel.navAlvo)) {
                store.setItem(AUTO_ESTADO, rel.precisaPreencher ? 'preenchendo_tempomedio' : ('coletando_' + key));
            }
        }
    }

    // fila: array de chaves ('juntadas'|'retorno'|'tempomedio') na ordem a executar.
    // periodoTM: id do período pré-configurado para a data inicial do Tempo Médio.
    function iniciarAutomacao(fila, periodoTM) {
        fila = (fila || []).filter(k => relatorioPorChave(k));
        if (!fila.length) { alert('Selecione ao menos um relatório para automatizar.'); return; }
        store.setItem('projudi_auto_fila', JSON.stringify(fila));
        store.setItem('projudi_auto_periodo_tm', periodoTM || '3a');
        const primeiro = relatorioPorChave(fila[0]);
        store.setItem(AUTO_ESTADO, primeiro.precisaPreencher ? 'preenchendo_tempomedio' : ('coletando_' + primeiro.key));
        store.setItem('projudi_auto_lock', String(Date.now()));
        atualizarPainel();
        setTimeout(() => navegarMenu(primeiro.navAlvo), 300);
    }

    function limparTudoAutomacao() {
        REPORTS_AUTOMACAO.forEach(({ cfg: c }) => {
            const n = parseInt(store.getItem(c.prefixo + 'num_paginas') || '0', 10);
            for (let i = 0; i < n; i++) store.removeItem(c.prefixo + 'pagina_' + i);
            store.removeItem(c.prefixo + 'num_paginas');
            store.removeItem(c.prefixo + 'rodando');
            store.removeItem(c.prefixo + 'ts');
        });
        store.removeItem(AUTO_ESTADO);
        store.removeItem('projudi_auto_fila');
        store.removeItem('projudi_auto_periodo_tm');
        store.removeItem('projudi_tempomedio_auto_iniciar');
        atualizarPainel();
    }

    function baixarPDFConjunto() {
        const secoes = REPORTS_AUTOMACAO
            .map(r => ({ dados: lerDadosDe(r.cfg.prefixo), cfg: r.cfg }))
            .filter(s => s.dados.length);
        if (!secoes.length) { alert('Nenhum dado coletado ainda.'); return; }
        try { gerarPDFConjunto(secoes); }
        catch (err) { alert('Erro ao gerar PDF conjunto: ' + err.message); console.error(err); }
    }

    // Painel flutuante (na página que tem o menu principal / página inicial)
    function atualizarPainel() {
        const painel = document.getElementById('painel-automacao');
        if (!painel) return;
        const estado = store.getItem(AUTO_ESTADO) || 'inativo';
        const contagens = REPORTS_AUTOMACAO.map(r => ({ r, n: lerDadosDe(r.cfg.prefixo).length }));
        const total = contagens.reduce((s, c) => s + c.n, 0);
        const emCurso = estado !== 'inativo' && estado !== 'concluido';
        const rotuloEstado = (() => {
            if (estado === 'inativo') return 'pronto';
            if (estado === 'concluido') return 'concluído';
            if (estado === 'ir_fim') return 'finalizando…';
            if (estado === 'preenchendo_tempomedio') return 'preenchendo filtros do tempo médio…';
            if (estado.startsWith('coletando_')) { const rel = relatorioPorChave(estado.slice(10)); return `coletando ${rel ? rel.rotulo.toLowerCase() : estado}…`; }
            if (estado.startsWith('ir_')) { const rel = relatorioPorChave(estado.slice(3)); return `indo para ${rel ? rel.rotulo.toLowerCase() : estado}…`; }
            return estado;
        })();

        const detalhe = contagens.map(c => `${c.r.rotulo}: ${c.n}`).join('  •  ');
        painel.querySelector('.pa-status').textContent = `Estado: ${rotuloEstado}  •  ${detalhe}`;
        painel.querySelector('#pa-iniciar').disabled = emCurso;
        painel.querySelector('#pa-pdf').disabled = total === 0;
        painel.querySelector('#pa-limpar').disabled = emCurso;
        painel.querySelectorAll('.pa-check, #pa-periodo-tm, #pa-marcar-tudo, #pa-desmarcar-tudo').forEach(el => { el.disabled = emCurso; });
    }

    function injetarPainel() {
        if (document.getElementById('painel-automacao')) return;
        // Só na página que hospeda o menu principal (evita duplicar em outros frames)
        if (!document.querySelector('#main-menu')) return;
        // Não injeta sobre a tela de resultados de um relatório
        if (detectarConfig()) return;

        const painel = document.createElement('div');
        painel.id = 'painel-automacao';
        const opcoesPeriodo = PERIODOS_TEMPOMEDIO.map(p => `<option value="${p.id}"${p.id === '3a' ? ' selected' : ''}>${p.rotulo}</option>`).join('');
        const linhasCheckbox = REPORTS_AUTOMACAO.map(r => `
                    <label class="pa-item">
                        <input type="checkbox" class="pa-check" data-key="${r.key}" checked> ${r.rotulo}
                        ${r.key === 'tempomedio' ? `<select id="pa-periodo-tm" class="projudi-select" title="Período usado como data inicial">${opcoesPeriodo}</select>` : ''}
                    </label>`).join('');
        painel.innerHTML = `
            <div class="pa-header">
                <span class="pa-titulo">Automação de relatórios</span>
                <div class="pa-controles">
                    <button class="pa-btn-colapsar" type="button" title="Recolher">▲</button>
                    <button class="pa-btn-fechar" type="button" title="Fechar">✕</button>
                </div>
            </div>
            <div class="pa-body">
                <div class="pa-status">—</div>
                <div class="pa-selecao">${linhasCheckbox}</div>
                <div class="pa-marcar">
                    <button id="pa-marcar-tudo" class="pa-link-btn" type="button">Marcar tudo</button>
                    <button id="pa-desmarcar-tudo" class="pa-link-btn" type="button">Desmarcar tudo</button>
                </div>
                <div class="pa-botoes">
                    <button id="pa-iniciar" class="projudi-btn" type="button" title="Extrai os relatórios marcados automaticamente">▶ Automatizar</button>
                    <button id="pa-pdf" class="projudi-btn" type="button" title="Gera um PDF único com os relatórios já coletados">⬇ PDF conjunto</button>
                    <button id="pa-limpar" class="projudi-btn" type="button" title="Apaga os dados acumulados de todos os relatórios">Limpar</button>
                </div>
                <div class="pa-dica">Rode em cada Atuação para acumular várias competências.</div>
            </div>`;
        document.body.appendChild(painel);
        painel.querySelector('#pa-iniciar').onclick = () => {
            const fila = [...painel.querySelectorAll('.pa-check:checked')].map(c => c.dataset.key);
            const periodoTM = painel.querySelector('#pa-periodo-tm').value;
            iniciarAutomacao(fila, periodoTM);
        };
        painel.querySelector('#pa-pdf').onclick = baixarPDFConjunto;
        painel.querySelector('#pa-limpar').onclick = limparTudoAutomacao;
        painel.querySelector('#pa-marcar-tudo').onclick = () => painel.querySelectorAll('.pa-check').forEach(c => { c.checked = true; });
        painel.querySelector('#pa-desmarcar-tudo').onclick = () => painel.querySelectorAll('.pa-check').forEach(c => { c.checked = false; });
        painel.querySelector('.pa-btn-colapsar').onclick = () => {
            const body = painel.querySelector('.pa-body');
            const btn = painel.querySelector('.pa-btn-colapsar');
            const recolhido = body.style.display === 'none';
            body.style.display = recolhido ? '' : 'none';
            btn.textContent = recolhido ? '▲' : '▼';
        };
        painel.querySelector('.pa-btn-fechar').onclick = () => painel.remove();
        atualizarPainel();
        // Mantém o painel vivo e conduz a navegação da automação (a coleta ocorre no
        // frame de conteúdo, que pode não recarregar este frame) — poll leve.
        setInterval(() => { atualizarPainel(); passoAutomacao(); }, 2000);
    }

    GM_addStyle(`
        #painel-automacao {
            position: fixed; top: 8px; right: 8px; z-index: 999999;
            background: #f7f7f2; border: 1px solid #63735f; border-radius: 6px;
            padding: 8px 10px; box-shadow: 0 2px 8px rgba(0,0,0,.25);
            font-family: Verdana, Arial, sans-serif; width: 300px;
        }
        #painel-automacao .pa-header { display: flex; align-items: center; justify-content: space-between; margin-bottom: 4px; }
        #painel-automacao .pa-titulo { font-weight: bold; font-size: .8em; color: #2d3748; }
        #painel-automacao .pa-controles { display: flex; gap: 3px; }
        #painel-automacao .pa-btn-colapsar, #painel-automacao .pa-btn-fechar {
            background: none; border: 1px solid #aaa; border-radius: 3px; cursor: pointer;
            font-size: .7em; padding: 1px 5px; color: #555; line-height: 1.3;
        }
        #painel-automacao .pa-btn-colapsar:hover, #painel-automacao .pa-btn-fechar:hover { background: #ddd; }
        #painel-automacao .pa-status { font-size: .72em; color: #444; margin-bottom: 6px; }
        #painel-automacao .pa-selecao { display: flex; flex-direction: column; gap: 3px; margin-bottom: 4px; }
        #painel-automacao .pa-item { font-size: .74em; color: #333; display: flex; align-items: center; gap: 4px; }
        #painel-automacao .pa-item input[type="checkbox"] { margin: 0; }
        #painel-automacao .pa-item .projudi-select { margin-left: 4px; padding: 1px 2px; font-size: 1em; }
        #painel-automacao .pa-marcar { display: flex; gap: 8px; margin-bottom: 6px; }
        #painel-automacao .pa-link-btn {
            background: none; border: none; padding: 0; cursor: pointer;
            font-size: .7em; color: #34556b; text-decoration: underline;
        }
        #painel-automacao .pa-link-btn:disabled { color: #999; cursor: not-allowed; }
        #painel-automacao .pa-botoes { display: flex; gap: 4px; flex-wrap: wrap; }
        #painel-automacao .projudi-btn { margin-left: 0; }
        #painel-automacao #pa-iniciar { background-color: #1e6b1e; border-color: #145214; }
        #painel-automacao #pa-pdf { background-color: #34556b; border-color: #26404f; }
        #painel-automacao #pa-limpar { background-color: #8a3b3b; border-color: #6e2f2f; }
        #painel-automacao .pa-dica { font-size: .66em; color: #777; margin-top: 5px; }
    `);

    function bootstrap() {
        console.log(`[Projudi] bootstrap — URL: ${location.href}`);
        injetarBotoes();       // botões nos relatórios (buttonBar)
        injetarPainel();       // painel de automação (página com o menu)
        passoAutomacao();      // avança a automação se estiver em estado de navegação
    }

    if (document.readyState === 'loading') {
        document.addEventListener('DOMContentLoaded', bootstrap);
    } else {
        bootstrap();
    }
})();
