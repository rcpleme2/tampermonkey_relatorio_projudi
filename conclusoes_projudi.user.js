// ==UserScript==
// @name         Exportar Conclusões Projudi para Excel
// @namespace    https://projudi2.tjpr.jus.br/
// @version      20.3
// @description  Coleta conclusões/retorno/juntadas/tempo médio/paralisados/remessas, exporta Excel ou PDF, e automatiza a extração conjunta a partir da página inicial
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

    // ── Estatísticas Gerais (processos ativos por atuação) ──────────────────────
    // Lê "Processos Ativos > Eletrônicos: N" da página inicial. Só existe nessa página
    // (fora dela retorna null, sem custo real — chamado sempre no bootstrap).
    function lerProcessosAtivos() {
        const labels = document.querySelectorAll('td.label label');
        for (const label of labels) {
            if (!/eletr[ôo]nicos\s*:/i.test(label.textContent)) continue;
            const tr = label.closest('tr');
            const tds = tr ? tr.querySelectorAll('td') : [];
            const txt = tds[1] ? tds[1].textContent : '';
            const m = /([\d.]+)/.exec(txt);
            if (m) { const n = parseInt(m[1].replace(/\./g, ''), 10); if (!isNaN(n)) return n; }
        }
        return null;
    }

    function lerMapaAtivos() {
        try { return JSON.parse(store.getItem('projudi_estatisticas_ativos') || '{}'); } catch (e) { return {}; }
    }

    // Grava/atualiza a contagem de processos ativos da atuação atual — acumulado entre
    // rodadas da automação (uma por atuação), como os demais relatórios.
    function gravarProcessosAtivosSeDisponivel() {
        const n = lerProcessosAtivos();
        if (n == null) return;
        const atuacao = lerAtuacao() || '(sem atuação)';
        const mapa = lerMapaAtivos();
        mapa[atuacao] = n;
        store.setItem('projudi_estatisticas_ativos', JSON.stringify(mapa));
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

    // Extrai o login do usuário do mesmo texto da célula "Dt. Análise Cartório" (ex.:
    // "16/07/2026" + "12849252930.est" em linhas separadas) — usado para agrupar o Tempo
    // Médio por usuário do cartório, já que o Projudi não expõe essa coluna separada.
    // Antes procurava o padrão "dígitos.sufixo" em td.textContent inteiro — mas a célula
    // pode ter OUTRO token nesse mesmo formato em algum elemento anterior (ex.: um ícone
    // de ajuda escondido), e o regex sempre casava com o primeiro que aparecesse — na
    // prática, sempre um token com o mesmo sufixo fixo (".tec"), nunca o login de verdade
    // quando o sufixo era outro (".anl", ".est" etc.). Em vez de casar por padrão em
    // qualquer lugar da célula, pega pela POSIÇÃO: a linha depois do <br> — a data fica
    // na primeira linha, o login sempre na última.
    function usuarioDeTexto(td) {
        if (!td) return '';
        const partes = (td.innerHTML || '').split(/<br\s*\/?>/i);
        if (partes.length < 2) return '';
        const ultimaLinha = norm(partes[partes.length - 1].replace(/<[^>]+>/g, ' '));
        const m = /(\d{5,}\.\w+)/.exec(ultimaLinha);
        return m ? m[1] : ultimaLinha;
    }

    // Competência = o nome completo da Atuação (ex.: "Vara de Família de Curitiba")
    function competenciaDe(atuacao) {
        return (atuacao || '').trim();
    }

    // Processo prioritário: o número recebe uma classe de cor extra (attentionBlue,
    // attentionPurple, attentionRed, ...); o normal fica apenas com "attention" (preto).
    // Memoiza o fallback de cor computada por className: getComputedStyle força um
    // recálculo de estilo (reflow) a cada chamada, e em páginas com 500 linhas quase
    // todas compartilham a mesma classe "normal" — sem cache isso travava o navegador
    // (centenas de reflows síncronos ao coletar uma página só).
    const _corPrioritarioCache = new Map();
    function emPrioritario(em) {
        if (!em) return false;
        const cls = em.className || '';
        // Verifica classes de cor (attentionRed, attentionBlue, attentionPurple …)
        if (cls.split(/\s+/).some(c => /^attention[A-Z]/.test(c) && c !== 'attentionBold')) return true;
        // Fallback: cor computada — vermelho indica processo prioritário (relatório tempoMédio)
        const chave = cls || '(sem-classe)';
        if (_corPrioritarioCache.has(chave)) return _corPrioritarioCache.get(chave);
        let resultado = false;
        try {
            const cor = window.getComputedStyle(em).color;
            const m = /rgb\((\d+),\s*(\d+),\s*(\d+)\)/.exec(cor);
            if (m && +m[1] > 150 && +m[2] < 100 && +m[3] < 100) resultado = true;
        } catch (e) {}
        _corPrioritarioCache.set(chave, resultado);
        return resultado;
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
        cabecalhos: ['Atuação', 'Dt. Conclusão', 'Processo', 'Classe', 'Seq.', 'Tipo de Conclusão', 'Privativa', 'Magistrado(a)', 'Pré-análise', 'Dt. Pré-análise', 'Agrupador', 'Prioritário'],
        larguras: [{ wch: 24 }, { wch: 18 }, { wch: 26 }, { wch: 18 }, { wch: 6 }, { wch: 30 }, { wch: 10 }, { wch: 30 }, { wch: 14 }, { wch: 14 }, { wch: 20 }, { wch: 11 }],
        extrai: (tds, atuacao) => {
            const pc = processoEclasse(tds[2]);
            const preAnaliseTexto = textoCelula(tds[7]);
            // A célula de Pré-análise traz "<matrícula>.asr - <nome do assessor>
            // (dd/mm/aaaa)" — a data de realização fica sempre entre parênteses, logo
            // após o nome. Extraímos separadamente para calcular há quanto tempo a
            // pré-análise ocorreu. Se não houver essa data entre parênteses (célula
            // vazia ou formato inesperado), preAnaliseData fica vazio e o tempo
            // decorrido é exibido como "—" no PDF (sem quebrar nada).
            const mData = /\((\d{2}\/\d{2}\/\d{4})\)/.exec(preAnaliseTexto);
            return {
                atuacao,
                competencia: competenciaDe(atuacao),
                dtRemessa: textoCelula(tds[1]),
                processo: pc.processo,
                classe: pc.classe,
                seq: textoCelula(tds[3]),
                tipoConclusao: textoCelula(tds[4]),
                privativa: textoCelula(tds[5]),
                responsavel: textoCelula(tds[6]),
                preAnalise: preAnaliseTexto,
                preAnaliseData: mData ? mData[1] : '',
                agrupador: textoCelula(tds[8]),
                prioritario: emPrioritario(tds[2].querySelector('em')),
            };
        },
        linha: (d) => [d.atuacao, d.dtRemessa, d.processo, d.classe, d.seq, d.tipoConclusao, d.privativa, d.responsavel, d.preAnalise, d.preAnaliseData, d.agrupador, d.prioritario ? 'Sim' : 'Não'],
        pdf: {
            titulo: 'Conclusões',
            atosTitulo: 'Conclusões pendentes',
            agingTitulo: 'Conclusões por tempo de espera',
            tabelaTitulo: 'Tabela discriminada das conclusões pendentes',
            dataCampo: 'dtRemessa',
            dataTitulo: 'Remessa para conclusão mais antiga',
            processoCampo: 'processo',
            tipoCampo: 'tipoConclusao',
            mediaLabel: 'conclusões / dia',
            distribuicoes: [
                { titulo: 'Conclusões por Magistrado(a)', campo: 'responsavel', topN: 12 },
                { titulo: 'Conclusões por Agrupador', campo: 'agrupador', topN: 12 },
                { titulo: 'Conclusões por Tipo de Conclusão', campo: 'tipoConclusao', topN: 12 },
            ],
            // Larguras somam 174mm — cabem na área útil (~186mm em A4 retrato com margem
            // de 12mm); antes somavam 206mm e a tabela vazava a borda direita da página.
            colunas: [
                { header: 'Atuação', width: 22, get: (d) => d.atuacao },
                { header: 'Processo', width: 26, get: (d) => d.processo },
                { header: 'Classe', width: 16, get: (d) => d.classe },
                { header: 'Tipo de Conclusão', width: 24, get: (d) => d.tipoConclusao },
                { header: 'Magistrado(a)', width: 26, get: (d) => d.responsavel },
                { header: 'Agrupador', width: 18, get: (d) => d.agrupador },
                { header: 'Dt. Conclusão', width: 18, get: (d) => d.dtRemessa },
                { header: 'Dias', width: 8, get: (d, ctx) => diasDecorridos(d.dtRemessa, ctx.now) },
                // Faltava no relatório geral (só existia no PDF por Juiz) — mesma lógica de
                // diasDesdePreAnalise usada lá (data extraída entre parênteses na célula de
                // Pré-análise).
                { header: 'Dias desde pré-análise', width: 16, get: (d, ctx) => diasDesdePreAnalise(d, ctx.now) || '—' },
            ],
            // Conclusões prioritárias primeiro; dentro de cada grupo (prioritárias/normais),
            // da data de conclusão mais antiga para a mais nova (ver montarTabelaGenerico).
            ordenarPrioritarioPrimeiro: true,
        },
        pdfPorJuiz: true, // habilita o botão extra "PDF por Juiz" (ver injetarBotoes)
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
            agingTitulo: 'Tempo de pendência da juntada',
            tabelaTitulo: 'Tabela discriminada das juntadas pendentes',
            dataCampo: 'dataEnvio',
            dataTitulo: 'Juntada pendente mais antiga',
            processoCampo: 'processo',
            tipoCampo: 'tipoDocumento',
            distribuicoes: [
                { titulo: 'Pendências por Função', campo: 'funcao', topN: 10 },
                // Os três gráficos abaixo vão para a 2ª página do resumo (pagina2), com
                // mais itens (15) já que ganham a página inteira só para eles.
                { titulo: 'Juntadas pendentes por pessoa', campo: 'juntadoPor', topN: 15, pagina2: true },
                // Largura total (span 2) para caber o nome completo do tipo de documento
                { titulo: 'Processos por Tipo de Documento', campo: 'tipoDocumento', topN: 15, span: 2, limpar: (s) => s.replace(/^juntada de\s+/i, ''), pagina2: true },
                // Ranking (sem "Outros" — cada processo é único, não faz sentido agregar o
                // restante). minValor: 2 — só processos com MAIS DE UMA juntada pendente;
                // se nenhum se qualificar, o gráfico inteiro é omitido (ver montarResumoGenerico).
                { titulo: 'Processos com mais de uma juntada pendente (15 maiores)', campo: 'processo', topN: 15, span: 2, semOutros: true, minValor: 2, pagina2: true },
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
        cabecalhos: ['Processo', 'Dt. Análise', 'Dt. Análise Cartório', 'Usuário Cartório', 'Tipo de Conclusão', 'Classe Processual', 'Dias p/ Cumprimento', 'Prioritário'],
        larguras: [{ wch: 26 }, { wch: 16 }, { wch: 20 }, { wch: 18 }, { wch: 26 }, { wch: 40 }, { wch: 18 }, { wch: 11 }],
        extrai: (tds, atuacao) => {
            const emProc = tds[1].querySelector('em');
            const processo = emProc ? emProc.textContent.trim() : textoCelula(tds[1]);
            const dtAnalise = dataDeTexto(tds[3]);
            const dtCartorio = dataDeTexto(tds[4]);
            const usuarioCartorio = usuarioDeTexto(tds[4]); // login abaixo da data, na mesma célula
            const classeCompleta = textoAteBr(tds[6]) || textoCelula(tds[6]);
            const classe = classeCompleta.split(' (')[0].trim(); // classe sem o "(Assunto Principal)"
            // Dias para cumprimento = Dt. Análise Cartório - Dt. Análise
            const tA = parseDataBR(dtAnalise), tC = parseDataBR(dtCartorio);
            const dias = (tA != null && tC != null) ? Math.max(0, Math.round((tC - tA) / DIA_MS)) : null;
            return {
                processo,
                dtAnalise,
                dtCartorio,
                usuarioCartorio,
                tipoConclusao: textoAteBr(tds[5]),
                classe,
                dias,
                prioritario: emPrioritario(emProc),
                atuacao: atuacao || '',
                competencia: competenciaDe(atuacao),
            };
        },
        linha: (d) => [d.processo, d.dtAnalise, d.dtCartorio, d.usuarioCartorio, d.tipoConclusao, d.classe,
                       (d.dias == null ? '' : String(d.dias)), d.prioritario ? 'Sim' : 'Não'],
        pdfCustom: (dados, somenteResumo) => gerarPDFTempoMedio(dados, somenteResumo),
    };

    // "Processos Paralisados" e "Remessas em Aberto" usam a MESMA tabela/URL
    // (processoBuscaParalisado.do) — o que muda é o filtro "opcaoBusca" selecionado no
    // formulário (1 = Na secretaria, 3 = Em remessa exceto conclusos), então o cabeçalho
    // sozinho não basta para distinguir os dois; é preciso ler o rádio marcado.
    function opcaoBuscaParalisadoSelecionada() {
        const el = document.querySelector('input[name="opcaoBusca"]:checked');
        return el ? el.value : null;
    }

    // Relatório de Processos Paralisados (processoBuscaParalisado.do, opcaoBusca=1 "Na
    // secretaria"). Colunas da tabela: [0]semáforo [1]checkbox [2]Processo [3]Seq.
    // [4]Classe Processual [5]Dias Paralisado [6]Razão Externa (não usado) [7]Último Movimento
    const CFG_PARALISADOS = {
        prefixo: 'projudi_paralisados_',
        detecta: (cab) => /dias\s+paralisado/i.test(cab) && opcaoBuscaParalisadoSelecionada() !== '3',
        minTds: 8,
        usaAtuacao: false,
        nomeArquivo: 'paralisados_projudi',
        rotulos: { coletar: 'Extrair Paralisados', coletarMais: 'Extrair mais (Paralisados)', baixar: '⬇ Baixar Paralisados' },
        cabecalhos: ['Processo', 'Seq.', 'Classe Processual', 'Dias Paralisado', 'Último Movimento', 'Prioritário'],
        larguras: [{ wch: 26 }, { wch: 10 }, { wch: 40 }, { wch: 16 }, { wch: 50 }, { wch: 11 }],
        extrai: (tds, atuacao) => {
            const emProc = tds[2].querySelector('em');
            const processo = emProc ? emProc.textContent.trim() : textoCelula(tds[2]);
            const diasTexto = textoCelula(tds[5]);
            const dias = /^\d+$/.test(diasTexto) ? parseInt(diasTexto, 10) : null;
            return {
                processo,
                seq: textoCelula(tds[3]),
                classe: textoCelula(tds[4]),
                dias,
                ultimoMovimento: textoCelula(tds[7]),
                prioritario: emPrioritario(emProc),
                atuacao: atuacao || '',
                competencia: competenciaDe(atuacao),
            };
        },
        linha: (d) => [d.processo, d.seq, d.classe, (d.dias == null ? '' : String(d.dias)), d.ultimoMovimento, d.prioritario ? 'Sim' : 'Não'],
        pdfCustom: (dados, somenteResumo) => gerarPDFParalisados(dados, somenteResumo),
    };

    // Relatório de Remessas em Aberto — mesma tabela do Paralisados, mas com o filtro
    // "Em remessa, exceto processos conclusos" (opcaoBusca=3). Extrai só o que foi pedido:
    // Processo, Classe Processual, Dias Paralisado e Último Movimento (sem Seq./Razão Externa).
    const CFG_REMESSAS = {
        prefixo: 'projudi_remessas_',
        detecta: (cab) => /dias\s+paralisado/i.test(cab) && opcaoBuscaParalisadoSelecionada() === '3',
        minTds: 8,
        usaAtuacao: false,
        nomeArquivo: 'remessas_abertas_projudi',
        rotulos: { coletar: 'Extrair Remessas', coletarMais: 'Extrair mais (Remessas)', baixar: '⬇ Baixar Remessas' },
        cabecalhos: ['Processo', 'Classe Processual', 'Dias Paralisado', 'Último Movimento', 'Prioritário'],
        larguras: [{ wch: 26 }, { wch: 40 }, { wch: 16 }, { wch: 50 }, { wch: 11 }],
        extrai: (tds, atuacao) => {
            const emProc = tds[2].querySelector('em');
            const processo = emProc ? emProc.textContent.trim() : textoCelula(tds[2]);
            const diasTexto = textoCelula(tds[5]);
            const dias = /^\d+$/.test(diasTexto) ? parseInt(diasTexto, 10) : null;
            return {
                processo,
                classe: textoCelula(tds[4]),
                dias,
                ultimoMovimento: textoCelula(tds[7]),
                prioritario: emPrioritario(emProc),
                atuacao: atuacao || '',
                competencia: competenciaDe(atuacao),
            };
        },
        linha: (d) => [d.processo, d.classe, (d.dias == null ? '' : String(d.dias)), d.ultimoMovimento, d.prioritario ? 'Sim' : 'Não'],
        pdfCustom: (dados, somenteResumo) => gerarPDFRemessas(dados, somenteResumo),
    };

    // Relatório de Processos Suspensos por Prazo Indeterminado (processoBuscaSuspenso.do,
    // acessado pelo número na página inicial). Colunas da tabela: [0]Processo
    // [1]Classe Processual [2]Prazo [3]Início Suspensão [4]Fim Suspensão [5]Dias Paralisado.
    // Sem pdf/pdfCustom próprio — entra na seção "Estatísticas Gerais" do relatório
    // conjunto (ver gerarPDFConjunto), não como relatório individual com resumo/gráficos.
    const CFG_SUSPENSOS = {
        prefixo: 'projudi_suspensos_',
        // Aparece no Relatório PDF mesmo com zero pendentes, desde que já coletado (ver
        // foiColetado) — "zero suspensos" é uma informação relevante, não um vazio a
        // esconder.
        mostrarSeVazio: true,
        detecta: (cab) => /in[íi]cio\s+suspens[ãa]o/i.test(cab),
        minTds: 6,
        usaAtuacao: false,
        nomeArquivo: 'suspensos_indeterminado_projudi',
        rotulos: { coletar: 'Extrair Suspensos', coletarMais: 'Extrair mais (Suspensos)', baixar: '⬇ Baixar Suspensos' },
        cabecalhos: ['Processo', 'Classe Processual', 'Início Suspensão', 'Dias Paralisado', 'Prioritário'],
        larguras: [{ wch: 26 }, { wch: 30 }, { wch: 16 }, { wch: 16 }, { wch: 11 }],
        extrai: (tds, atuacao) => {
            const emProc = tds[0].querySelector('em');
            const processo = emProc ? emProc.textContent.trim() : textoCelula(tds[0]);
            const diasTexto = textoCelula(tds[5]);
            const dias = /^\d+$/.test(diasTexto) ? parseInt(diasTexto, 10) : null;
            return {
                processo,
                classe: textoCelula(tds[1]),
                inicioSuspensao: textoCelula(tds[3]),
                dias,
                prioritario: emPrioritario(emProc),
                atuacao: atuacao || '',
                competencia: competenciaDe(atuacao),
            };
        },
        linha: (d) => [d.processo, d.classe, d.inicioSuspensao, (d.dias == null ? '' : String(d.dias)), d.prioritario ? 'Sim' : 'Não'],
        // Tratado como mais uma tarefa do Cartório (mesmo esquema genérico de Juntadas/
        // Retorno — ver gerarPDFConjunto): entra no resumo/tabela por seção e na mini-
        // tabela de situação da capa, sem precisar de uma página "Estatísticas Gerais"
        // separada.
        pdf: {
            titulo: 'Suspensos por Prazo Indeterminado',
            atosTitulo: 'Processos suspensos',
            agingTitulo: 'Tempo de suspensão',
            tabelaTitulo: 'Tabela discriminada dos processos suspensos por prazo indeterminado',
            dataCampo: 'inicioSuspensao',
            dataTitulo: 'Suspensão mais antiga',
            processoCampo: 'processo',
            tipoCampo: 'classe',
            distribuicoes: [
                { titulo: 'Suspensos por Classe Processual', campo: 'classe', topN: 12 },
            ],
            colunas: [
                { header: 'Processo', width: 30, get: (d) => d.processo },
                { header: 'Classe', width: 40, get: (d) => d.classe },
                { header: 'Início Suspensão', width: 24, get: (d) => d.inicioSuspensao },
                { header: 'Dias Paralisado', width: 20, get: (d) => (d.dias == null ? '' : String(d.dias)) },
            ],
        },
    };

    // Verifica qual "Situação" está marcada no formulário de Audiências (audienciaForm) —
    // a tela mostra o form e a table.resultTable juntos, sempre com as mesmas colunas
    // (Para Hoje/Pendentes/Movimentadas Hoje/Para a data compartilham o cabeçalho), então
    // só dá para saber qual filtro está ativo lendo o rádio marcado.
    function situacaoAudienciaSelecionada() {
        const el = document.querySelector('input[name="idSituacaoAudiencia"]:checked');
        return el ? el.value : null;
    }

    // Primeiro relatório específico da categoria Crime: Audiências Pendentes
    // (audiencia/busca.do, idSituacaoAudiencia=8 "Pendentes"). Colunas da tabela:
    // [0]Processo/Recurso [1]Partes [2]Local da Audiência [3]Data [4]Tipo da Audiência
    // [5]Modalidade [6]Situação da Audiência.
    const CFG_AUDIENCIAS = {
        prefixo: 'projudi_audiencias_',
        // Mostra o card mesmo com zero audiências pendentes, desde que já coletado (ver
        // foiColetado) — "nenhuma audiência pendente" é a informação, não um vazio.
        mostrarSeVazio: true,
        detecta: (cab) => /tipo\s+da\s+audi[êe]ncia/i.test(cab) && /situa[çc][ãa]o\s+da\s+audi[êe]ncia/i.test(cab)
            && situacaoAudienciaSelecionada() === '8',
        minTds: 7,
        usaAtuacao: false,
        nomeArquivo: 'audiencias_pendentes_projudi',
        rotulos: { coletar: 'Extrair Audiências', coletarMais: 'Extrair mais (Audiências)', baixar: '⬇ Baixar Audiências' },
        cabecalhos: ['Processo', 'Tipo da Audiência', 'Data da Audiência', 'Dias até a Audiência', 'Prioritário'],
        larguras: [{ wch: 26 }, { wch: 34 }, { wch: 18 }, { wch: 18 }, { wch: 11 }],
        extrai: (tds, atuacao) => {
            const emProc = tds[0].querySelector('em');
            const processo = emProc ? emProc.textContent.trim() : textoCelula(tds[0]);
            return {
                processo,
                dataAudiencia: textoCelula(tds[3]),
                tipoAudiencia: textoCelula(tds[4]),
                prioritario: emPrioritario(emProc),
                atuacao: atuacao || '',
                competencia: competenciaDe(atuacao),
            };
        },
        // "Dias até a Audiência" é sempre recalculado na hora de exportar (Excel ou PDF),
        // nunca gravado na coleta — a diferença é entre a data da audiência e o momento em
        // que o relatório é gerado, não o momento em que os dados foram coletados.
        linha: (d) => [d.processo, d.tipoAudiencia, d.dataAudiencia, fmtDiferencaDias(diasAteAudiencia(d.dataAudiencia, Date.now())), d.prioritario ? 'Sim' : 'Não'],
        pdfCustom: (dados, somenteResumo) => gerarPDFAudiencias(dados, somenteResumo),
    };

    // Diferença de dias entre a data da audiência e "now" (positivo = audiência no
    // futuro, negativo = já deveria ter ocorrido). null quando a data não pôde ser lida.
    function diasAteAudiencia(dataAudiencia, now) {
        const t = parseDataBR(dataAudiencia);
        return t == null ? null : Math.round((t - now) / DIA_MS);
    }
    function fmtDiferencaDias(dias) { return dias == null ? '' : String(dias); }

    // ── Audiências Designadas (Crime) — audiencia/pautaAudiencia.do, "Ver Pauta de
    // Horários" ───────────────────────────────────────────────────────────────────
    // Relatório de resumo + tabela discriminada (guardada dentro do resumo — ver
    // secaoTemTabela): total de audiências
    // designadas, o último dia com audiência e quantos processos distintos têm audiência
    // nesse dia (contados um a um, não pela coluna "Agendadas" — ver expandirLinhasDoDia),
    // a data da audiência mais distante por tipo, e uma situação (verde/amarelo/vermelho)
    // conforme o quão longe no futuro a agenda já está lotada.
    //
    // "Audiência de Instrução", "Audiência de Instrução e Julgamento" e "Audiência de
    // Interrogatório" contam como um tipo só ("Audiência de Instrução") — pedido do
    // usuário. A observação sobre essa junção aparece no relatório (ver OBSERVACAO_TIPOS_AD
    // em montarResumoAudienciasDesignadas).
    function normalizarTipoAudiencia(tipo) {
        const t = (tipo || '').trim();
        if (/^audi[êe]ncia\s+de\s+instru[çc][ãa]o(\s+e\s+julgamento)?$/i.test(t)) return 'Audiência de Instrução';
        if (/^audi[êe]ncia\s+de\s+interrogat[óo]rio$/i.test(t)) return 'Audiência de Instrução';
        return t;
    }
    const OBSERVACAO_TIPOS_AD = 'Observação: "Audiência de Instrução", "Audiência de Instrução e Julgamento" e '
        + '"Audiência de Interrogatório" foram tratadas como um único tipo ("Audiência de Instrução").';

    const CFG_AUDIENCIAS_DESIGNADAS = {
        prefixo: 'projudi_audienciasdesignadas_',
        mostrarSeVazio: true,
        // Tem tabela discriminada (data/hora/processo/tipo), mas guardada dentro do
        // resumo (resumo.tabela) em vez de uma lista de registros como os demais
        // relatórios — ver secaoTemTabela em gerarPDFConjunto e montarTabelaAudienciasDesignadas.
        detecta: () => false, // nunca detectado por cabeçalho de tabela — só entra via automação (ver navegarMenu/injetarBotoes)
        usaAtuacao: false,
        nomeArquivo: 'audiencias_designadas_projudi',
        pdfCustom: (dados) => gerarPDFAudienciasDesignadas(dados),
    };

    // Tela de "Ver Pauta de Horários" (audiencia/pautaAudiencia.do) — form de filtros +
    // resultado paginado por dia/local, tudo numa única página (o Projudi não pagina esse
    // relatório: mesmo pedindo 10 anos, só volta as datas que realmente têm audiência).
    function formularioPautaAudiencias() {
        const form = document.getElementById('pautaAudienciaForm');
        return form && form.querySelector('#divTiposAudiencia') ? form : null;
    }

    // Marca todos os tipos de audiência, define a Data Fim para 10 anos à frente de hoje e
    // pesquisa. Não depende do "Todas" cascatear os checkboxes individuais via JS da
    // página — marca cada input[name="idsTiposAudiencia"] diretamente, garantindo o mesmo
    // resultado independente desse comportamento.
    function preencherEPesquisarPautaAudiencias() {
        const form = formularioPautaAudiencias();
        if (!form) return;

        // O Projudi cascateia os tipos individuais a partir de um CLIQUE de verdade no
        // checkbox "Todas" (onclick da própria página) — só marcar checked=true e disparar
        // "change" não é suficiente (regressão: parou de funcionar depois de um merge, que
        // silenciosamente perdeu o .click() real). Clica de verdade e, como rede de
        // segurança, ainda confere se sobrou algum tipo desmarcado.
        const checkTodos = form.querySelector('input[name="checkMarcaTodos"]');
        if (checkTodos && !checkTodos.checked) {
            console.log('[Projudi Audiências Designadas] clicando no checkbox "Todas"');
            checkTodos.click();
        }
        const todos = form.querySelectorAll('input[name="idsTiposAudiencia"]');
        const semMarcar = [...todos].filter(chk => !chk.checked);
        if (semMarcar.length) {
            console.warn(`[Projudi Audiências Designadas] "Todas" não marcou ${semMarcar.length} tipo(s) — marcando manualmente`);
            semMarcar.forEach(chk => { chk.checked = true; chk.dispatchEvent(new Event('change', { bubbles: true })); });
        }
        console.log(`[Projudi Audiências Designadas] ${todos.length} tipo(s) de audiência marcado(s) (Todas=${checkTodos ? checkTodos.checked : 'n/d'})`);

        const campoFim = form.querySelector('#dataFinal');
        if (campoFim) {
            const daqui10anos = new Date();
            daqui10anos.setFullYear(daqui10anos.getFullYear() + 10);
            campoFim.value = formatarDataBR(daqui10anos);
            campoFim.dispatchEvent(new Event('input', { bubbles: true }));
            campoFim.dispatchEvent(new Event('change', { bubbles: true }));
        }
        console.log(`[Projudi Audiências Designadas] Data Fim="${campoFim ? campoFim.value : '?'}"`);

        const btn = document.getElementById('searchButton') || form.querySelector('input[type="button"][name="button"]');
        console.log(`[Projudi Audiências Designadas] botão de pesquisa encontrado=${!!btn}; clicando em 1,5s`);
        setTimeout(() => {
            if (btn && !btn.disabled) btn.click(); else form.submit();
        }, 1500);
    }

    // Detecta se a página já tem os resultados da pauta (pelo menos uma tabela interna
    // com o cabeçalho Horário/Tipo da Audiência) — a tela mostra form + resultado juntos
    // desde o primeiro carregamento, então "existe table.resultTable" sozinho não basta.
    function pautaAudienciasTemResultados() {
        return [...document.querySelectorAll('table.resultTable')].some(t => {
            const thead = t.querySelector(':scope > thead');
            const cab = thead ? thead.textContent : '';
            return /hor[áa]rio/i.test(cab) && /tipo\s+da\s+audi[êe]ncia/i.test(cab);
        });
    }

    // Lê todas as linhas (uma por horário/tipo) das tabelas internas de cada dia — a
    // "Pauta de Horários" agrupa por Local + Data (tabela externa) e, dentro de cada
    // grupo, uma tabela interna própria com colunas Horário/Modalidade/Criadas/Agendadas/
    // Pauta Auto./Via Grade/Tipo da Audiência. A linha oculta de detalhe (id="rowN", com o
    // <div class="extendedinfo"> onde os processos aparecem ao expandir) é ignorada aqui.
    function lerLinhasPautaAudiencias() {
        const linhas = [];
        document.querySelectorAll('table.resultTable').forEach(tabela => {
            const thead = tabela.querySelector(':scope > thead');
            if (!thead) return;
            const cab = thead.textContent;
            if (!/hor[áa]rio/i.test(cab) || !/tipo\s+da\s+audi[êe]ncia/i.test(cab)) return; // só as tabelas internas (por dia)

            const trExterna = tabela.closest('tr');
            const linkData = trExterna ? trExterna.querySelector('td a.link') : null;
            const data = linkData ? linkData.textContent.trim() : '';
            if (!data) return;

            tabela.querySelectorAll(':scope > tbody > tr').forEach(tr => {
                if (/^row\d+$/.test(tr.id)) return; // linha oculta de detalhe (extendedinfo)
                const tds = tr.querySelectorAll(':scope > td');
                if (tds.length < 8) return;
                linhas.push({
                    data,
                    horario: textoCelula(tds[1]),
                    criadas: parseInt(textoCelula(tds[3]), 10) || 0,
                    agendadas: parseInt(textoCelula(tds[4]), 10) || 0,
                    tipoAudiencia: normalizarTipoAudiencia(textoCelula(tds[7])),
                    elLinha: tr,
                });
            });
        });
        return linhas;
    }

    // Expande um LOTE de linhas de uma vez (clica em todos os ícones "+" do lote antes de
    // esperar) — bem mais rápido que expandir uma a uma, já que a espera é a mesma AJAX
    // independente de quantas expansões estão em voo ao mesmo tempo. O tamanho do lote é
    // um meio-termo: grande o bastante para acelerar de verdade, pequeno o bastante para
    // não sobrecarregar o Projudi nem estourar o tempo de resposta de uma leva só.
    const TAMANHO_LOTE_EXPANSAO_AD = 15;

    async function expandirLoteDeLinhas(lote) {
        lote.forEach(linha => {
            const link = linha.elLinha.querySelector('a[class^="linkProcessos"]');
            linha._temLink = !!link;
            if (link) link.click();
        });
        // Espera única para o lote inteiro — a base de 1,2s cobre uma resposta normal do
        // Projudi; o acréscimo por linha dá uma folga extra quando o lote é grande (mais
        // requisições AJAX disputando o servidor ao mesmo tempo).
        const espera = 1200 + lote.length * 80;
        await new Promise(r => setTimeout(r, espera));
        lote.forEach(linha => {
            if (!linha._temLink) { linha.processos = []; return; }
            const detalhe = linha.elLinha.nextElementSibling;
            const texto = detalhe ? detalhe.textContent : '';
            linha.processos = texto.match(/\d{7}-\d{2}\.\d{4}\.\d\.\d{2}\.\d{4}/g) || [];
        });
    }

    // Progresso da expansão, exposto para o painel flutuante mostrar (ver atualizarPainel)
    // — como a leitura pode demorar bastante em pautas grandes, sem isso o usuário via só
    // "Coletando Audiências Designadas…" parado por vários segundos, sem saber se travou.
    const CHAVE_PROGRESSO_AD = 'projudi_audienciasdesignadas_progresso';
    function atualizarProgressoExpansaoAD(processados, total) {
        store.setItem(CHAVE_PROGRESSO_AD, JSON.stringify({ processados, total }));
        atualizarPainel();
    }
    function lerProgressoExpansaoAD() {
        try { return JSON.parse(store.getItem(CHAVE_PROGRESSO_AD) || 'null'); }
        catch (e) { return null; }
    }

    // Expande TODAS as linhas lidas, em lotes (ver TAMANHO_LOTE_EXPANSAO_AD), preenchendo
    // linha.processos. Necessário para a tabela discriminada completa (data/hora/processo/
    // tipo), além de alimentar a contagem do último dia, a última audiência por tipo e a
    // concentração por dia da semana — todas baseadas nos processos de verdade, não na
    // coluna "Agendadas".
    async function expandirTodasAsLinhas(linhas) {
        const total = linhas.length;
        store.removeItem(CHAVE_PROGRESSO_AD);
        atualizarProgressoExpansaoAD(0, total);
        for (let i = 0; i < total; i += TAMANHO_LOTE_EXPANSAO_AD) {
            const lote = linhas.slice(i, i + TAMANHO_LOTE_EXPANSAO_AD);
            await expandirLoteDeLinhas(lote);
            const processados = Math.min(i + lote.length, total);
            const semProcesso = lote.filter(l => l.processos.length === 0).length;
            if (semProcesso) {
                console.warn(`[Projudi Audiências Designadas] ${semProcesso} linha(s) do lote ${Math.floor(i / TAMANHO_LOTE_EXPANSAO_AD) + 1} sem processo vinculado — serão ignoradas (ver filtro de linhas sem processo).`);
            }
            console.log(`[Projudi Audiências Designadas] expandindo processos: ${processados}/${total} linha(s)`);
            atualizarProgressoExpansaoAD(processados, total);
        }
    }

    // Dias úteis (segunda a sexta — Date#getDay(): 0=domingo...6=sábado) em que cada
    // processo (não cada linha — uma linha pode ter vários processos) tem audiência,
    // agrupados por tipo. Fins de semana são descartados (pedido do usuário: só dias
    // úteis) — o Projudi não deveria pautar audiência em fim de semana, mas por segurança
    // essas entradas (se existirem) não entram na análise.
    const ROTULOS_DIA_UTIL = { 1: 'Segunda', 2: 'Terça', 3: 'Quarta', 4: 'Quinta', 5: 'Sexta' };
    function concentracaoPorDiaSemana(linhasComProcesso) {
        const porDia = new Map(); // diaSemana(1-5) -> Map(tipo -> quantidade de processos)
        linhasComProcesso.forEach(l => {
            const ts = parseDataBR(l.data);
            if (ts == null) return;
            const dow = new Date(ts).getDay();
            if (dow < 1 || dow > 5) return;
            const porTipoDia = porDia.get(dow) || new Map();
            porTipoDia.set(l.tipoAudiencia, (porTipoDia.get(l.tipoAudiencia) || 0) + l.processos.length);
            porDia.set(dow, porTipoDia);
        });
        return [1, 2, 3, 4, 5].map(dow => {
            const porTipoDia = porDia.get(dow) || new Map();
            const porTipo = [...porTipoDia.entries()].map(([tipo, quantidade]) => ({ tipo, quantidade })).sort((a, b) => b.quantidade - a.quantidade);
            return { diaSemana: ROTULOS_DIA_UTIL[dow], total: porTipo.reduce((s, it) => s + it.quantidade, 0), porTipo };
        });
    }

    // Conduz a extração inteira (síncrona + a expansão assíncrona de todas as linhas em
    // lotes) e grava o resumo calculado no mesmo formato de armazenamento usado por
    // lerDadosDe/criarColetor (uma "página" com um array de 1 item), para participar do
    // Relatório PDF/automação como qualquer outro relatório.
    async function coletarAudienciasDesignadas() {
        console.log('[Projudi Audiências Designadas] iniciando extração da Pauta de Horários');
        const todasAsLinhas = lerLinhasPautaAudiencias();
        console.log(`[Projudi Audiências Designadas] ${todasAsLinhas.length} linha(s) lida(s) — iniciando expansão dos processos em lotes de ${TAMANHO_LOTE_EXPANSAO_AD}`);

        await expandirTodasAsLinhas(todasAsLinhas);

        // Só entram na análise as linhas com pelo menos um processo vinculado (pedido do
        // usuário) — uma linha da pauta sem processo reconhecido ao expandir é ignorada em
        // TODOS os cálculos abaixo, não só na tabela.
        const linhas = todasAsLinhas.filter(l => l.processos.length > 0);
        const ignoradas = todasAsLinhas.length - linhas.length;
        if (ignoradas > 0) console.log(`[Projudi Audiências Designadas] ${ignoradas} linha(s) sem processo vinculado foram ignoradas`);

        // Total de audiências designadas = total de processos com audiência marcada (uma
        // linha pode ter mais de um processo no mesmo horário) — não mais a soma da coluna
        // "Agendadas", já que agora só contam linhas com processo de verdade vinculado.
        const totalDesignadas = linhas.reduce((s, l) => s + l.processos.length, 0);

        let ultimaData = null, ultimaDataTs = -Infinity;
        linhas.forEach(l => {
            const ts = parseDataBR(l.data);
            if (ts != null && ts > ultimaDataTs) { ultimaDataTs = ts; ultimaData = l.data; }
        });

        // Processos distintos no último dia — contados um a um a partir da expansão, não
        // pela soma da coluna "Agendadas" (pedido do usuário).
        let processosUltimoDia = [];
        if (ultimaData) {
            const set = new Set();
            linhas.filter(l => l.data === ultimaData).forEach(l => l.processos.forEach(p => set.add(p)));
            processosUltimoDia = [...set];
        }
        const totalProcessosUltimoDia = processosUltimoDia.length;

        // Data mais distante por tipo (já com "Audiência de Instrução e Julgamento" somada
        // em "Audiência de Instrução" — ver normalizarTipoAudiencia), com o(s) número(s) de
        // processo daquela audiência mais distante.
        const maxTsPorTipo = new Map();
        linhas.forEach(l => {
            const ts = parseDataBR(l.data);
            if (ts == null) return;
            const atual = maxTsPorTipo.get(l.tipoAudiencia);
            if (atual == null || ts > atual) maxTsPorTipo.set(l.tipoAudiencia, ts);
        });
        const porTipo = [...maxTsPorTipo.entries()].map(([tipo, ts]) => {
            const linhasDoTipoNaData = linhas.filter(l => l.tipoAudiencia === tipo && parseDataBR(l.data) === ts);
            const processos = [...new Set(linhasDoTipoNaData.flatMap(l => l.processos))];
            return { tipo, data: linhasDoTipoNaData[0].data, processos };
        }).sort((a, b) => b.data.localeCompare(a.data));

        // Tabela discriminada completa: data/hora/processo/tipo, uma linha por processo
        // (uma linha da pauta pode ter mais de um processo agendado no mesmo horário).
        const tabela = [];
        linhas.forEach(l => {
            l.processos.forEach(p => tabela.push({ data: l.data, horario: l.horario, processo: p, tipoAudiencia: l.tipoAudiencia }));
        });
        tabela.sort((a, b) => {
            const ta = parseDataBR(a.data) || 0, tb = parseDataBR(b.data) || 0;
            return ta - tb || a.horario.localeCompare(b.horario);
        });

        const concentracaoDiaSemana = concentracaoPorDiaSemana(linhas);

        const resumo = {
            geradoEm: new Date().toISOString(),
            totalDesignadas,
            ultimaData,
            processosUltimoDia,
            totalProcessosUltimoDia,
            porTipo,
            tabela,
            concentracaoDiaSemana,
            competencia: competenciaDe(lerAtuacao()),
        };
        console.log(`[Projudi Audiências Designadas] resumo calculado: totalDesignadas=${totalDesignadas} ultimaData=${ultimaData} totalProcessosUltimoDia=${totalProcessosUltimoDia} tipos=${porTipo.length} linhasTabela=${tabela.length} ignoradas(sem processo)=${ignoradas}`);

        const prefixo = CFG_AUDIENCIAS_DESIGNADAS.prefixo;
        store.setItem(prefixo + 'pagina_0', JSON.stringify([resumo]));
        store.setItem(prefixo + 'num_paginas', '1');
        store.setItem(prefixo + 'coletado', '1');
        store.removeItem(CHAVE_PROGRESSO_AD);

        avancarAutomacao(CFG_AUDIENCIAS_DESIGNADAS);
    }

    // ── Audiências Realizadas (Crime) — audiencia/estatistica.do, "Audiências na
    // Vara" ──────────────────────────────────────────────────────────────────────
    // Diferente dos demais relatórios de Crime: uma única tela de filtros/resultado
    // (estatisticaAudienciaForm) que devolve só TOTAIS (não uma lista de processos), e o
    // detalhamento por usuário exige uma pesquisa PARA CADA usuário do combo — não dá pra
    // ler tudo de uma vez. Por isso a coleta é uma fila (mesmo padrão do Tempo Médio "mês a
    // mês"): primeiro uma pesquisa "geral" (usuário em branco) pro total do período, depois
    // uma pesquisa por usuário, na sequência, acumulando os totais.
    const CFG_AUDIENCIAS_REALIZADAS = {
        prefixo: 'projudi_audienciasrealizadas_',
        mostrarSeVazio: true,
        detecta: () => false, // nunca detectado por cabeçalho — só entra via automação (ver navegarMenu/injetarBotoes)
        usaAtuacao: false,
        nomeArquivo: 'audiencias_realizadas_projudi',
        pdfCustom: (dados) => gerarPDFAudienciasRealizadas(dados),
    };

    // Quantidade mínima de audiências realizadas para um usuário aparecer no detalhamento
    // (pedido do usuário) — os demais ainda contam para o total geral, só não entram na
    // lista por usuário.
    const MIN_AUDIENCIAS_POR_USUARIO_AR = 10;

    function formularioAudienciasRealizadas() {
        const form = document.getElementById('estatisticaAudienciaForm');
        return form && form.querySelector('#usuario') ? form : null;
    }

    // Data final = último dia do mês anterior ao vigente; data inicial = 3 anos antes dessa
    // data (pedido do usuário) — sempre um período de 3 anos "fechado" (nunca inclui o mês
    // corrente, ainda em andamento).
    function periodoAudienciasRealizadas(agora) {
        const fim = new Date(agora.getFullYear(), agora.getMonth(), 0);
        const inicio = new Date(fim.getFullYear() - 3, fim.getMonth(), fim.getDate());
        return { dataInicio: formatarDataBR(inicio), dataFim: formatarDataBR(fim) };
    }

    function preencherPeriodoAR(form, periodo) {
        const campoIni = form.querySelector('#dataInicio');
        const campoFim = form.querySelector('#dataFim');
        [[campoIni, periodo.dataInicio], [campoFim, periodo.dataFim]].forEach(([campo, valor]) => {
            if (!campo) return;
            campo.value = valor;
            campo.dispatchEvent(new Event('input', { bubbles: true }));
            campo.dispatchEvent(new Event('change', { bubbles: true }));
        });
        // "Dt. Agendamento" (tipoBuscaDataAudiencia=1) — o Projudi às vezes já marca isso
        // por padrão, mas não sempre; marcar explicitamente evita depender desse padrão.
        const radioAgendamento = form.querySelector('input[name="tipoBuscaDataAudiencia"][value="1"]');
        if (radioAgendamento) {
            radioAgendamento.checked = true;
            radioAgendamento.dispatchEvent(new Event('click', { bubbles: true }));
            radioAgendamento.dispatchEvent(new Event('change', { bubbles: true }));
        }
    }

    function setUsuarioAR(form, valor) {
        const sel = form.querySelector('#usuario');
        if (!sel) return;
        sel.value = valor;
        sel.dispatchEvent(new Event('change', { bubbles: true }));
    }

    // Lê as opções reais do combo de usuários (ignora o placeholder "-- CLIQUE AQUI PARA
    // SELECIONAR --", que tem value vazio) — feito UMA vez, na primeira pesquisa (geral),
    // já que o combo é o mesmo em toda pesquisa desta tela.
    function lerOpcoesUsuarioAR(form) {
        const sel = form.querySelector('#usuario');
        if (!sel) return [];
        return [...sel.querySelectorAll('option')]
            .filter(o => o.value)
            .map(o => ({ value: o.value, label: o.textContent.trim() }));
    }

    // Lê o valor (coluna Quantidade) da linha da tabela de resultado cujo rótulo bate com
    // "rotuloRegex" (ex.: "Total Realizadas", "Canceladas") — 0 se a linha não aparecer.
    function lerValorLinhaAR(rotuloRegex) {
        let valor = 0;
        document.querySelectorAll('table.resultTable tbody tr').forEach(tr => {
            const tds = tr.querySelectorAll(':scope > td');
            if (tds.length < 2) return;
            if (rotuloRegex.test(textoCelula(tds[0]))) {
                valor = parseInt(textoCelula(tds[1]).replace(/\D/g, ''), 10) || 0;
            }
        });
        return valor;
    }

    // Total de audiências realizadas — linha "Total Realizadas" (em negrito). Mesma tabela
    // serve tanto pra pesquisa geral quanto pra cada pesquisa por usuário.
    function lerTotalRealizadasAR() { return lerValorLinhaAR(/^total\s+realizadas$/i); }

    // Canceladas/Não Realizadas/Redesignadas — lidas tanto na pesquisa geral (total do
    // período) quanto em cada pesquisa por usuário (pedido do usuário: também por
    // magistrado), sempre da mesma tabela de resultado.
    function lerExtrasAR() {
        return {
            canceladas: lerValorLinhaAR(/^canceladas$/i),
            naoRealizadas: lerValorLinhaAR(/^n[ãa]o\s+realizadas$/i),
            redesignadas: lerValorLinhaAR(/^redesignadas$/i),
        };
    }

    function submitAR(form) {
        const btn = form.querySelector('#searchButton');
        setTimeout(() => { if (btn && !btn.disabled) btn.click(); else form.submit(); }, 1200);
    }

    const CHAVE_AGUARDANDO_AR = 'projudi_audienciasrealizadas_aguardando';
    const CHAVE_TOTAL_GERAL_AR = 'projudi_audienciasrealizadas_totalgeral';
    const CHAVE_EXTRAS_GERAL_AR = 'projudi_audienciasrealizadas_extrasgeral';
    const CHAVE_FILA_USUARIOS_AR = 'projudi_audienciasrealizadas_fila_usuarios';
    const CHAVE_ACUMULADO_AR = 'projudi_audienciasrealizadas_acumulado';
    const CHAVE_PROGRESSO_AR = 'projudi_audienciasrealizadas_progresso';
    const CHAVE_TOTAL_USUARIOS_AR = 'projudi_audienciasrealizadas_total_usuarios';

    function atualizarProgressoAR(processados, total) {
        store.setItem(CHAVE_PROGRESSO_AR, JSON.stringify({ processados, total }));
        atualizarPainel();
    }

    function lerProgressoAR() {
        try { return JSON.parse(store.getItem(CHAVE_PROGRESSO_AR) || 'null'); }
        catch (e) { return null; }
    }

    // Primeira entrada nesta automação: preenche o período e pesquisa com usuário em branco
    // (todos) — o total dessa pesquisa é o total geral do período; o combo lido nessa mesma
    // tela vira a fila de usuários a percorrer depois.
    function iniciarBuscaAudienciasRealizadas() {
        const form = formularioAudienciasRealizadas();
        if (!form) return;
        const periodo = periodoAudienciasRealizadas(new Date());
        console.log(`[Projudi Audiências Realizadas] iniciando — período ${periodo.dataInicio} a ${periodo.dataFim}, pesquisa geral (todos os usuários)`);
        preencherPeriodoAR(form, periodo);
        setUsuarioAR(form, '');
        store.setItem(CHAVE_AGUARDANDO_AR, JSON.stringify({ tipo: 'geral' }));
        store.setItem('projudi_audienciasrealizadas_periodo', JSON.stringify(periodo));
        store.setItem(AUTO_ESTADO, 'coletando_audienciasrealizadas');
        submitAR(form);
    }

    // Avança para o próximo usuário da fila (ou finaliza, se a fila acabou).
    function avancarUsuarioAR(form) {
        // desembrulharArray, não JSON.parse cru — o valor às vezes volta JSON-codificado em
        // camadas (visto em outros relatórios, ver CHAVE_FILA_MESES_TM/lerDadosDe); sem
        // isso, um JSON.parse único podia deixar "fila" como STRING, e fila.shift() quebrava.
        const fila = desembrulharArray(store.getItem(CHAVE_FILA_USUARIOS_AR)) || [];
        if (!fila.length) { finalizarAudienciasRealizadas(); return; }
        const totalUsuarios = parseInt(store.getItem(CHAVE_TOTAL_USUARIOS_AR) || '0', 10);
        const prox = fila.shift();
        store.setItem(CHAVE_FILA_USUARIOS_AR, JSON.stringify(fila));
        console.log(`[Projudi Audiências Realizadas] pesquisando usuário "${prox.label}" (${totalUsuarios - fila.length}/${totalUsuarios})`);
        atualizarProgressoAR(totalUsuarios - fila.length - 1, totalUsuarios);
        setUsuarioAR(form, prox.value);
        store.setItem(CHAVE_AGUARDANDO_AR, JSON.stringify({ tipo: 'usuario', value: prox.value, label: prox.label }));
        submitAR(form);
    }

    // Chamado quando a tela recarrega já com resultado — decide o que aquele resultado
    // significa (geral ou de qual usuário, ver CHAVE_AGUARDANDO_AR) e segue a fila.
    function processarResultadoAudienciasRealizadas() {
        const form = formularioAudienciasRealizadas();
        if (!form) return;
        let aguardando = null;
        try { aguardando = JSON.parse(store.getItem(CHAVE_AGUARDANDO_AR) || 'null'); } catch (e) { /* ignore */ }
        const total = lerTotalRealizadasAR();

        if (!aguardando || aguardando.tipo === 'geral') {
            store.setItem(CHAVE_TOTAL_GERAL_AR, String(total));
            const extras = lerExtrasAR();
            store.setItem(CHAVE_EXTRAS_GERAL_AR, JSON.stringify(extras));
            const usuarios = lerOpcoesUsuarioAR(form);
            console.log(`[Projudi Audiências Realizadas] total geral do período: ${total} (canceladas=${extras.canceladas} não realizadas=${extras.naoRealizadas} redesignadas=${extras.redesignadas}) — ${usuarios.length} usuário(s) a percorrer`);
            store.setItem(CHAVE_FILA_USUARIOS_AR, JSON.stringify(usuarios));
            store.setItem(CHAVE_TOTAL_USUARIOS_AR, String(usuarios.length));
            store.setItem(CHAVE_ACUMULADO_AR, JSON.stringify([]));
            atualizarProgressoAR(0, usuarios.length);
            avancarUsuarioAR(form);
            return;
        }

        const extrasUsuario = lerExtrasAR();
        console.log(`[Projudi Audiências Realizadas] "${aguardando.label}": ${total} realizada(s), ${extrasUsuario.canceladas} cancelada(s), ${extrasUsuario.naoRealizadas} não realizada(s), ${extrasUsuario.redesignadas} redesignada(s)`);
        const acumulado = desembrulharArray(store.getItem(CHAVE_ACUMULADO_AR)) || [];
        acumulado.push({
            usuario: aguardando.value, nome: aguardando.label, quantidade: total,
            canceladas: extrasUsuario.canceladas, naoRealizadas: extrasUsuario.naoRealizadas, redesignadas: extrasUsuario.redesignadas,
        });
        store.setItem(CHAVE_ACUMULADO_AR, JSON.stringify(acumulado));
        avancarUsuarioAR(form);
    }

    function limparEstadoTransitorioAR() {
        [CHAVE_AGUARDANDO_AR, CHAVE_TOTAL_GERAL_AR, CHAVE_EXTRAS_GERAL_AR, CHAVE_FILA_USUARIOS_AR,
         CHAVE_ACUMULADO_AR, CHAVE_PROGRESSO_AR, CHAVE_TOTAL_USUARIOS_AR, 'projudi_audienciasrealizadas_periodo']
            .forEach(k => store.removeItem(k));
    }

    // Fila de usuários vazia: monta o resumo final (total geral + detalhamento por usuário,
    // já filtrado pelo mínimo — ver MIN_AUDIENCIAS_POR_USUARIO_AR) e grava no mesmo formato
    // usado por lerDadosDe/criarColetor (uma "página" com um array de 1 item).
    function finalizarAudienciasRealizadas() {
        const totalGeral = parseInt(store.getItem(CHAVE_TOTAL_GERAL_AR) || '0', 10);
        const acumulado = desembrulharArray(store.getItem(CHAVE_ACUMULADO_AR)) || [];
        let periodo = { dataInicio: '', dataFim: '' };
        try { periodo = JSON.parse(store.getItem('projudi_audienciasrealizadas_periodo') || 'null') || periodo; } catch (e) { /* ignore */ }
        let extrasGeral = { canceladas: 0, naoRealizadas: 0, redesignadas: 0 };
        try { extrasGeral = JSON.parse(store.getItem(CHAVE_EXTRAS_GERAL_AR) || 'null') || extrasGeral; } catch (e) { /* ignore */ }

        const porUsuario = acumulado
            .filter(u => u.quantidade >= MIN_AUDIENCIAS_POR_USUARIO_AR)
            .sort((a, b) => b.quantidade - a.quantidade);
        const usuariosIgnorados = acumulado.length - porUsuario.length;

        const resumo = {
            geradoEm: new Date().toISOString(),
            totalGeral,
            canceladas: extrasGeral.canceladas,
            naoRealizadas: extrasGeral.naoRealizadas,
            redesignadas: extrasGeral.redesignadas,
            periodo,
            porUsuario,
            usuariosIgnorados,
            totalUsuarios: acumulado.length,
            competencia: competenciaDe(lerAtuacao()),
        };
        console.log(`[Projudi Audiências Realizadas] resumo calculado: totalGeral=${totalGeral} canceladas=${extrasGeral.canceladas} naoRealizadas=${extrasGeral.naoRealizadas} redesignadas=${extrasGeral.redesignadas} usuarios=${acumulado.length} exibidos(>=${MIN_AUDIENCIAS_POR_USUARIO_AR})=${porUsuario.length} ignorados=${usuariosIgnorados}`);

        const prefixo = CFG_AUDIENCIAS_REALIZADAS.prefixo;
        store.setItem(prefixo + 'pagina_0', JSON.stringify([resumo]));
        store.setItem(prefixo + 'num_paginas', '1');
        store.setItem(prefixo + 'coletado', '1');
        limparEstadoTransitorioAR();

        avancarAutomacao(CFG_AUDIENCIAS_REALIZADAS);
    }

    // ── Apreensões Pendentes (Crime) — processo/criminal/apreensao.do ───────────────
    // Página de filtros + resultado na MESMA tela (mesmo padrão de Paralisados/
    // Audiências): o Projudi já vem com "Motivo do encerramento = (Apreensão não
    // encerrada)" e "Status = ATIVO" marcados por padrão — que é exatamente o filtro de
    // "pendente" pedido pelo usuário — então não mexemos em nenhum desses selects, só
    // clicamos em Pesquisar.
    function colunasApreensoes(sufixoRotulo) {
        return {
            cabecalhos: ['Data do Registro', 'Número', 'Tipo da Apreensão', 'Processo', 'Classe Processual', 'Data de Encerramento', 'Motivo do Encerramento', 'Localização Interna', 'Descrição', 'Prioritário'],
            larguras: [{ wch: 14 }, { wch: 14 }, { wch: 26 }, { wch: 26 }, { wch: 24 }, { wch: 16 }, { wch: 20 }, { wch: 20 }, { wch: 34 }, { wch: 11 }],
            extrai: (tds, atuacao) => {
                const emProc = tds[3].querySelector('em');
                const processo = emProc ? emProc.textContent.trim() : textoCelula(tds[3]);
                const classeCompleta = textoAteBr(tds[4]) || textoCelula(tds[4]);
                const classe = classeCompleta.split(' (')[0].trim();
                return {
                    dataRegistro: textoCelula(tds[0]),
                    numero: textoCelula(tds[1]),
                    tipo: textoCelula(tds[2]),
                    processo,
                    classe,
                    dataEncerramento: textoCelula(tds[5]),
                    motivoEncerramento: textoCelula(tds[6]),
                    localizacao: textoCelula(tds[7]),
                    descricao: textoCelula(tds[8]),
                    prioritario: emPrioritario(emProc),
                    atuacao: atuacao || '',
                    competencia: competenciaDe(atuacao),
                };
            },
            linha: (d) => [d.dataRegistro, d.numero, d.tipo, d.processo, d.classe, d.dataEncerramento, d.motivoEncerramento, d.localizacao, d.descricao, d.prioritario ? 'Sim' : 'Não'],
        };
    }

    const COLS_APREENSOES = colunasApreensoes();

    const LISTA_TIPOS_APREENSAO = [
        'Armas de fogo', 'Munições', 'Explosivos', 'Objetos', 'Entorpecentes', 'Valores', 'Plantas', 'Animais',
        'Armas brancas', 'Aeronaves', 'Alimentos, bebidas, medicamentos e outros produtos perecíveis',
        'Ativos Financeiros, cheques e outros títulos de crédito',
        'Computadores, acessórios, insumos e outros produtos de informática', 'Documentos',
        'Eletroeletrônicos diversos', 'Embarcações', 'Equipamentos de Caça e Pesca (exceto armas)',
        'Material Biológico', 'Outros Bens Móveis', 'Outros Meios de Transporte',
        'Objetos Pessoais ou Domésticos', 'Produtos Florestais', 'Veículos Automotores', 'Bens Imóveis',
        'Pedras, Metais Preciosos, jóias, quadros, objetos de arte, objeto de coleção e antiguidade',
        'Acessórios de Armas de Fogo e Produtos Controlados pelo EB',
    ];

    const CFG_APREENSOES = {
        prefixo: 'projudi_apreensoes_',
        // Mostra a linha "Bens Apreendidos" mesmo com zero pendências, desde que já
        // coletado — "zero apreensões pendentes" é uma informação, não um vazio a esconder.
        mostrarSeVazio: true,
        detecta: (cab) => /tipo\s+da\s+apreens[ãa]o/i.test(cab),
        minTds: 9,
        usaAtuacao: false,
        nomeArquivo: 'apreensoes_pendentes_projudi',
        rotulos: { coletar: 'Extrair Apreensões', coletarMais: 'Extrair mais (Apreensões)', baixar: '⬇ Baixar Apreensões' },
        cabecalhos: COLS_APREENSOES.cabecalhos,
        larguras: COLS_APREENSOES.larguras,
        extrai: COLS_APREENSOES.extrai,
        linha: COLS_APREENSOES.linha,
        pdf: {
            titulo: 'Bens Apreendidos Pendentes',
            atosTitulo: 'Apreensões pendentes',
            agingTitulo: 'Apreensões por tempo de registro',
            tabelaTitulo: 'Tabela discriminada das apreensões pendentes',
            dataCampo: 'dataRegistro',
            dataTitulo: 'Registro de apreensão mais antigo',
            processoCampo: 'processo',
            tipoCampo: 'tipo',
            // Apreensões não distingue prioritário/normal (não faz sentido pra esse
            // relatório) — some com o KPI "Prioritários pendentes" e com a separação por
            // prioridade no gráfico de aging (vira barra única por faixa).
            semPrioridade: true,
            // Tabela discriminada dividida em subtabelas por Tipo da Apreensão, na mesma
            // ordem do <select name="idTipoApreensaoBusca"> do Projudi.
            agruparPor: 'tipo',
            ordemGrupos: LISTA_TIPOS_APREENSAO,
            distribuicoes: [
                { titulo: 'Apreensões por Tipo', campo: 'tipo', topN: 12 },
                { titulo: 'Apreensões por Localização Interna', campo: 'localizacao', topN: 10 },
                { titulo: 'Apreensões por Classe Processual', campo: 'classe', topN: 12 },
            ],
            colunas: [
                { header: 'Número', width: 18, get: (d) => d.numero },
                { header: 'Tipo', width: 28, get: (d) => d.tipo },
                { header: 'Processo', width: 30, get: (d) => d.processo },
                { header: 'Classe', width: 24, get: (d) => d.classe },
                { header: 'Dt. Registro', width: 20, get: (d) => d.dataRegistro },
                { header: 'Localização', width: 24, get: (d) => d.localizacao },
                { header: 'Descrição', width: 28, get: (d) => d.descricao },
            ],
        },
    };

    function formularioApreensoes() {
        const form = document.getElementById('apreensaoForm');
        return form && form.querySelector('#idMotivoEncerramentoApreensaoBusca') ? form : null;
    }

    // Força o motivo de encerramento para "(Apreensão não encerrada)" e clica em
    // Pesquisar — os demais filtros ficam no padrão da tela.
    function preencherEPesquisarApreensoes() {
        const form = formularioApreensoes();
        if (!form) return;

        const selMotivo = form.querySelector('#idMotivoEncerramentoApreensaoBusca');
        if (selMotivo && selMotivo.value !== '0') {
            selMotivo.value = '0';
            selMotivo.dispatchEvent(new Event('change', { bubbles: true }));
        }
        console.log(`[Projudi Apreensões] motivo de encerramento definido como "(Apreensão não encerrada)" (value=${selMotivo ? selMotivo.value : 'n/d'})`);

        const btn = document.getElementById('pesquisar') || form.querySelector('input[type="submit"]');
        console.log(`[Projudi Apreensões] botão de pesquisa encontrado=${!!btn}; clicando em 1,5s`);
        setTimeout(() => {
            console.log('[Projudi Apreensões] clicando em Pesquisar — o site pode demorar para responder, aguarde.');
            if (btn && !btn.disabled) btn.click(); else form.submit();

            setTimeout(() => {
                const aindaNoFormulario = !document.querySelector('table.resultTable');
                console.log(`[Projudi Apreensões] diagnóstico 15s depois — aindaSemResultado=${aindaNoFormulario}`);
                if (aindaNoFormulario) {
                    console.warn('[Projudi Apreensões] ainda sem resultado após 15s — o site pode estar lento; se persistir, clique em Pesquisar manualmente.');
                }
            }, 15000);
        }, 1500);
    }

    // ── Coletor genérico (parametrizado por configuração) ───────────────────────

    function criarColetor(cfg) {
        const KEY_RODANDO     = cfg.prefixo + 'rodando';
        const KEY_NUM_PAGINAS = cfg.prefixo + 'num_paginas';
        const KEY_PAGINA_PREF = cfg.prefixo + 'pagina_';
        const KEY_TS          = cfg.prefixo + 'ts';
        // Marca que a coleta chegou ao fim ao menos uma vez — diferente de KEY_NUM_PAGINAS
        // (que fica em 0 tanto quando nunca rodou quanto quando rodou e não achou nada),
        // isso permite ao PDF conjunto distinguir "não coletado" de "coletado, zero registros"
        // (ver Suspensos por Prazo Indeterminado em gerarPDFConjunto/baixarPDFConjunto).
        const KEY_COLETADO    = cfg.prefixo + 'coletado';

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
            store.removeItem(KEY_COLETADO);
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
                // ":scope > td" (só filhos diretos) — não "td" puro, que também pega tds de
                // QUALQUER tabela aninhada dentro de uma célula (ex.: a coluna "Partes" das
                // Audiências tem uma table.form própria por dentro), o que bagunçava a
                // contagem/índice dos tds da linha.
                const tds = tr.querySelectorAll(':scope > td');
                if (tds.length < cfg.minTds) return;
                dados.push(cfg.extrai(tds, atuacao));
            });
            console.log(`[Projudi] coletarPaginaAtual — ${linhas.length} linhas encontradas, ${dados.length} extraídas (minTds=${cfg.minTds})`);
            return dados;
        }

        // Tamanho de página alvo para o seletor "estatisticaPageSizeOptions" (só existe na
        // tela de resultados do Tempo Médio). 500 travava o site do Projudi — a tabela com
        // 1000 <tr> (500 processos + linhas de detalhe) era pesada demais para o próprio
        // Projudi renderizar/paginar. Com a busca agora feita mês a mês (bem menos
        // registros por pesquisa — ver preencherEPesquisarTempoMedio), 100 já é seguro.
        const TAMANHO_PAGINA_ALVO = '100';

        function iniciar() {
            // Antes de iniciar, ajusta a página para exibir TAMANHO_PAGINA_ALVO registros
            // por vez (se essa opção existir no seletor). Quando o valor do select muda, o
            // Projudi recarrega a página; o estado KEY_RODANDO já estará salvo e continuar()
            // será chamado automaticamente ao reiniciar (bloco rodando && !obsoleta).
            const sel = document.querySelector('select[name="estatisticaPageSizeOptions"]');
            if (sel) {
                const opcoesDisponiveis = [...sel.options].map(o => o.value);
                const alvo = opcoesDisponiveis.includes(TAMANHO_PAGINA_ALVO) ? TAMANHO_PAGINA_ALVO : null;
                if (alvo && sel.value !== alvo) {
                    console.log(`[Projudi] alterando pageSize de ${sel.value} para ${alvo} — aguardando reload`);
                    store.setItem(KEY_RODANDO, '1');
                    marcarAtividade();
                    sel.value = alvo;
                    sel.dispatchEvent(new Event('change', { bubbles: true }));
                    return;
                }
                if (!alvo) console.log(`[Projudi] opção ${TAMANHO_PAGINA_ALVO} não existe no seletor de tamanho de página — mantendo ${sel.value}`);
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
                store.setItem(KEY_COLETADO, '1');
                render();
                const extra = cfg.usaAtuacao ? ` de ${contarAtuacoes()} atuação(ões)` : '';
                const dica = cfg.usaAtuacao ? ' Troque de atuação e colete mais, ou baixe a planilha.' : ' Baixe a planilha ou colete mais.';
                atualizarStatus(`Coleta concluída. Acumulado: ${total} registros${extra}.${dica}`);
                // Tempo Médio busca mês a mês (ver preencherEPesquisarTempoMedio) — se ainda
                // restam meses na fila, volta para a tela de filtros e pesquisa o próximo em
                // vez de avançar para o próximo relatório da automação.
                if (cfg === CFG_TEMPOMEDIO && lerFilaMesesTempoMedio().length > 0) {
                    console.log('[Projudi TM] mês concluído — ainda restam meses na fila, buscando o próximo');
                    store.setItem(AUTO_ESTADO, 'ir_tempomedio');
                    setTimeout(passoAutomacao, 900);
                } else {
                    avancarAutomacao(cfg); // se a automação estiver ativa, segue para o próximo passo
                }
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

        function pdf(somenteResumo) {
            try {
                const dados = lerTudo();
                if (!dados.length) { atualizarStatus('Nenhum registro coletado para exportar.'); return; }
                if (cfg.pdfCustom) cfg.pdfCustom(dados, somenteResumo); else gerarPDF(dados, cfg, somenteResumo);
                // Após exportar o PDF, limpa os dados acumulados automaticamente — evita
                // que uma coleta antiga fique acumulada/misturada com a próxima.
                limparTudo();
                render();
                const extraResumo = somenteResumo ? ' (apenas resumo)' : '';
                atualizarStatus(`✓ PDF gerado com ${dados.length} registros${extraResumo}. Dados acumulados apagados — pronto para nova coleta.`);
            } catch (err) {
                atualizarStatus(`Erro ao gerar PDF: ${err.message}`);
                console.error('[Exportar Projudi]', err);
            }
        }

        // Só usado pelo relatório de Conclusões (cfg.pdfPorJuiz === true): um PDF com uma
        // seção por juiz responsável, para poder ser expedido individualmente.
        function pdfPorJuiz() {
            try {
                const dados = lerTudo();
                if (!dados.length) { atualizarStatus('Nenhum registro coletado para exportar.'); return; }
                gerarPDFConclusoesPorJuiz(dados);
                limparTudo();
                render();
                atualizarStatus(`✓ PDF por juiz gerado com ${dados.length} registros. Dados acumulados apagados — pronto para nova coleta.`);
            } catch (err) {
                atualizarStatus(`Erro ao gerar PDF por juiz: ${err.message}`);
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
            const bPdfJuiz = document.getElementById('btn-pdf-juiz');
            if (bPdfJuiz) { bPdfJuiz.disabled = total === 0; bPdfJuiz.textContent = `PDF por Juiz (${total})`; }
            bLimpar.disabled = total === 0;

            if (total > 0) {
                const extra = cfg.usaAtuacao ? ` de ${contarAtuacoes()} atuação(ões)` : '';
                atualizarStatus(`Acumulado: ${total} registros${extra}.`);
            } else {
                atualizarStatus('');
            }
        }

        return { iniciar, continuar, baixar, pdf, pdfPorJuiz, limpar, render, rodando, obsoleta,
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
    // Paleta sóbria (tons dessaturados, ar institucional) — mesmo papel semântico de
    // antes, mas com menos saturação/brilho para não parecer um dashboard "colorido".
    const COR = {
        tinta:    [26, 26, 26],     // títulos fortes
        tintaSec: [82, 81, 78],     // texto secundário
        muted:    [130, 128, 122],  // eixos / rótulos em maiúsculas
        grade:    [222, 221, 214],  // hairline / bordas
        base:     [195, 194, 183],  // linhas de base
        cartao:   [244, 244, 241],  // fundo de card (neutro, não azulado)
        azul:     [58, 90, 125],    // métrica principal / série "normal" (slate blue)
        aqua:     [82, 116, 103],   // secundário / positivo (verde-acinzentado)
        ambar:    [156, 116, 46],   // atenção / faixa intermediária (ocre)
        vermelho: [146, 58, 58],    // PRIORITÁRIO / crítico (terracota escuro)
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

    // semOutros: quando true, corta em topN sem somar o restante em "Outros" — usado em
    // rankings (ex.: top processos) onde o "Outros" agregado não faz sentido/domina o gráfico.
    function contarPorCampo(dados, campo, topN, limpar, semOutros, minValor) {
        const mapa = new Map();
        dados.forEach(d => {
            let k = (d[campo] || '').trim();
            if (limpar) k = limpar(k).trim();
            k = k || '(vazio)';
            mapa.set(k, (mapa.get(k) || 0) + 1);
        });
        let arr = [...mapa.entries()].map(([label, valor]) => ({ label, valor })).sort((a, b) => b.valor - a.valor);
        // minValor descarta contagens baixas antes do corte de topN (ex.: "processos com
        // mais de uma juntada pendente" não deve trazer processos com apenas uma).
        if (minValor != null) arr = arr.filter(i => i.valor >= minValor);
        if (arr.length > topN) {
            if (semOutros) {
                arr = arr.slice(0, topN);
            } else {
                const resto = arr.slice(topN).reduce((s, i) => s + i.valor, 0);
                arr = arr.slice(0, topN);
                arr.push({ label: 'Outros', valor: resto });
            }
        }
        return arr;
    }

    const COR_PRIORITARIO = COR.vermelho;  // realce dos prioritários (mesma cor em todo o relatório)

    function contarPrioritarios(dados) {
        return dados.reduce((n, d) => n + (d.prioritario ? 1 : 0), 0);
    }

    // ── Classificação de situação (Cartório/Gabinete) ───────────────────────────
    // Usada na página "Situação da Unidade" do relatório conjunto: dá um veredito
    // objetivo (não só números) sobre cada tarefa do Cartório e cada magistrado(a) do
    // Gabinete, pensado para leitura rápida numa correição.

    // Dias decorridos (número) a partir de uma data BR; null quando não há data válida.
    function diasNum(str, now) {
        const ts = parseDataBR(str);
        return ts == null ? null : Math.max(0, Math.floor((now - ts) / DIA_MS));
    }

    // Classifica pela pendência mais antiga (em dias), com limites próprios por domínio:
    // Cartório — regular até 30 dias, atenção de 31 a 90, crítico acima de 90; Gabinete —
    // regular até 30 dias, atenção de 31 a 120, crítico acima de 120 (ver limites em
    // gerarPDFConjunto).
    function classificarSituacaoPorDias(dias, limiteAtencao, limiteCritico) {
        if (dias == null) return 'regular';
        if (dias > limiteCritico) return 'critico';
        if (dias > limiteAtencao) return 'atencao';
        return 'regular';
    }

    const SITUACAO_INFO = {
        critico: { rotulo: 'Crítico', cor: COR.vermelho },
        atencao: { rotulo: 'Atenção', cor: COR.ambar },
        regular: { rotulo: 'Regular', cor: COR.aqua },
    };

    // A pior situação entre várias — usada para o veredito geral de um domínio (Cartório
    // ou Gabinete) a partir da situação de cada item que o compõe.
    function piorSituacao(situacoes) {
        if (situacoes.includes('critico')) return 'critico';
        if (situacoes.includes('atencao')) return 'atencao';
        return 'regular';
    }

    function maiorDias(itens) {
        return itens.reduce((max, it) => (it.dias != null && (max == null || it.dias > max) ? it.dias : max), null);
    }

    // Converte os dados de uma tarefa/relatório em itens {dias, prioritario} para
    // classificarSituacaoPorDias(). Paralisados/Remessas já trazem "dias" pronto (não uma data);
    // os demais calculam a partir do campo de data do próprio cfg.pdf.
    function itensParaClassificacao(dados, cfg, now) {
        if (cfg === CFG_PARALISADOS || cfg === CFG_REMESSAS || cfg === CFG_SUSPENSOS) {
            return dados.map(d => ({ dias: d.dias, prioritario: !!d.prioritario }));
        }
        const campo = cfg.pdf.dataCampo;
        return dados.map(d => ({ dias: diasNum(d[campo], now), prioritario: !!d.prioritario }));
    }

    // Lista as competências distintas presentes num conjunto de dados (ordem de
    // primeira ocorrência) — usado para indicar, no cabeçalho dos relatórios
    // individuais, a que atuação(ões) os dados se referem.
    function listaCompetencias(dados) {
        const vistas = new Set();
        dados.forEach(d => {
            const c = (d.competencia || '').trim();
            if (c) vistas.add(c);
        });
        return [...vistas];
    }

    // Frase "Competência: X" / "Competências: X, Y" para compor no subtítulo dos
    // relatórios — vazia quando não há dado de competência disponível.
    function fraseCompetencias(dados) {
        const comps = listaCompetencias(dados);
        if (!comps.length) return '';
        return `${comps.length > 1 ? 'Competências' : 'Competência'}: ${comps.join(', ')}`;
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
        doc.setFont('PublicSans', 'bold'); doc.setFontSize(10); doc.setTextColor(...COR.tinta);
        doc.text(texto, x, y);
        doc.setDrawColor(...acento); doc.setLineWidth(0.7);
        doc.line(x, y + 1.6, x + Math.min(w, doc.getTextWidth(texto) + 2), y + 1.6);
    }

    // Card de destaque (KPI), com uma barra de acento colorida à esquerda indicando o
    // papel semântico do dado (azul=principal, vermelho=prioritário, âmbar=atenção,
    // aqua=secundário). subs: array de linhas secundárias. Se central=true, o conteúdo
    // é centralizado horizontal e verticalmente dentro do card.
    // Encurta "texto" (com reticências) até caber em "larguraMax" com a fonte já ativa no
    // doc (setFont/setFontSize precisam ter sido chamados antes) — usado para o valor
    // principal do card nunca vazar pra fora da borda (bug relatado: um valor longo, como
    // vários números de processo, desenhado sem quebra ultrapassava o card).
    function textoTruncadoParaLargura(doc, texto, larguraMax) {
        const s = String(texto);
        if (doc.getTextWidth(s) <= larguraMax) return s;
        let t = s;
        while (t.length > 1 && doc.getTextWidth(t + '…') > larguraMax) t = t.slice(0, -1);
        return t + '…';
    }

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
            doc.setFont('PublicSans', 'bold'); doc.setFontSize(7.5); doc.setTextColor(...COR.muted);
            doc.text(String(titulo).toUpperCase(), cx, yy, { align: 'center' }); yy += 7;
            doc.setFont('PublicSans', 'bold'); doc.setFontSize(valor.length <= 26 ? 15 : 11); doc.setTextColor(...COR.tinta);
            doc.text(textoTruncadoParaLargura(doc, valor, w - 10), cx, yy, { align: 'center' }); yy += 5.5;
            doc.setFont('PublicSans', 'normal'); doc.setFontSize(8); doc.setTextColor(...COR.tintaSec);
            subs.forEach(s => { doc.text(doc.splitTextToSize(String(s), w - 10)[0], cx, yy, { align: 'center' }); yy += 4.2; });
            return;
        }

        const px = x + 5;
        doc.setFont('PublicSans', 'bold'); doc.setFontSize(7.5); doc.setTextColor(...COR.muted);
        doc.text(String(titulo).toUpperCase(), px, y + 6.5);
        const grande = valor.length <= 13;
        doc.setFont('PublicSans', 'bold'); doc.setFontSize(grande ? 20 : 14); doc.setTextColor(...COR.tinta);
        doc.text(textoTruncadoParaLargura(doc, valor, w - 8), px, y + (grande ? 16.5 : 15));
        if (subs.length) {
            doc.setFont('PublicSans', 'normal'); doc.setFontSize(8); doc.setTextColor(...COR.tintaSec);
            let yy = y + (grande ? 22 : 20);
            subs.forEach(s => { doc.text(doc.splitTextToSize(String(s), w - 8)[0], px, yy); yy += 4; });
        }
    }

    // Mede a altura que desenharCardLista precisaria para "valor" — chamar ANTES de
    // desenhar o card (a altura da caixa depende do texto, então precisa ser conhecida de
    // antemão), com a mesma fonte usada no desenho.
    function medirAlturaCardLista(doc, w, valor, temSub) {
        doc.setFont('PublicSans', 'bold'); doc.setFontSize(11);
        const linhas = doc.splitTextToSize(String(valor || '—'), w - 10);
        return 6.5 + linhas.length * 4.6 + (temSub ? 8 : 0) + 4;
    }

    // Variante de desenharCard para um "valor" que pode ser longo (ex.: vários números de
    // processo) — QUEBRA em várias linhas em vez de truncar com reticências, então nunca
    // corta informação (pedido do usuário). Altura fixa "h" — sempre calcular com
    // medirAlturaCardLista antes de chamar.
    function desenharCardLista(doc, x, y, w, h, titulo, valor, subLinha, acento) {
        acento = acento || COR.azul;
        doc.setDrawColor(...COR.grade); doc.setFillColor(...COR.cartao); doc.setLineWidth(0.2);
        doc.roundedRect(x, y, w, h, 2, 2, 'FD');
        doc.setFillColor(...acento);
        doc.roundedRect(x, y, 1.8, h, 0.9, 0.9, 'F');
        const px = x + 5;
        doc.setFont('PublicSans', 'bold'); doc.setFontSize(7.5); doc.setTextColor(...COR.muted);
        doc.text(String(titulo).toUpperCase(), px, y + 6.5);
        doc.setFont('PublicSans', 'bold'); doc.setFontSize(11); doc.setTextColor(...COR.tinta);
        const linhas = doc.splitTextToSize(String(valor || '—'), w - 10);
        doc.text(linhas, px, y + 12);
        if (subLinha) {
            doc.setFont('PublicSans', 'normal'); doc.setFontSize(8); doc.setTextColor(...COR.tintaSec);
            doc.text(subLinha, px, y + 12 + linhas.length * 4.6 + 3);
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
        const barH = Math.max(1.5, linhaH * 0.58);
        // A fonte acompanha a altura da linha: quando há muitos itens numa área pequena
        // (ex.: grades com várias seções na mesma página), reduz o texto em vez de
        // sobrepor uma linha na outra — nunca deixa o texto ilegível por colisão.
        const fonte = Math.max(5.5, Math.min(8, linhaH * 0.85));
        const offsetY = fonte * 0.15;

        doc.setFontSize(fonte);
        itens.forEach((it, i) => {
            const meio = topo + i * linhaH + linhaH / 2;
            doc.setFont('PublicSans', 'normal'); doc.setTextColor(...COR.tintaSec);
            doc.text(doc.splitTextToSize(it.label, rotuloW - 3)[0], x, meio + offsetY);
            const bw = Math.max(0.6, (it.valor / maxVal) * barMaxW);
            doc.setFillColor(...(it.cor || cor));
            doc.roundedRect(barX, meio - barH / 2, bw, barH, 0.6, 0.6, 'F');
            doc.setFont('PublicSans', 'bold'); doc.setTextColor(...COR.tinta);
            doc.text(fmt(it.valor), barX + bw + 2, meio + offsetY);
        });
    }

    // Cores de severidade por faixa (até 30d / 31-90d / mais de 90d), usadas como
    // marcador de ponto ao lado do rótulo — reforça a leitura sem depender só das barras.
    const COR_SEVERIDADE = [COR.aqua, COR.ambar, COR.vermelho];

    // Barras agrupadas por faixa: duas sub-barras (prioritários x normais) por linha —
    // o vermelho é sempre prioritário, nunca a faixa.
    function desenharBarrasFaixas(doc, x, y, w, h, titulo, faixas) {
        tituloSecao(doc, x, y + 4, w, titulo, COR.ambar);
        const legY = y + 10;
        doc.setFillColor(...COR_PRIORITARIO); doc.rect(x, legY - 2.4, 3, 3, 'F');
        doc.setFont('PublicSans', 'normal'); doc.setFontSize(7); doc.setTextColor(...COR.tintaSec);
        doc.text('Prioritários', x + 4, legY);
        const q2 = x + 28;
        doc.setFillColor(...COR.azul); doc.rect(q2, legY - 2.4, 3, 3, 'F');
        doc.text('Normais', q2 + 4, legY);

        const topo = y + 15;
        const areaH = Math.max(6, h - 15);
        const rotuloW = Math.min(36, w * 0.32);
        const valorW = 11;
        const barX = x + rotuloW;
        const barMaxW = Math.max(6, w - rotuloW - valorW);
        const maxVal = Math.max(1, ...faixas.map(f => Math.max(f.prioritarios, f.normais)));
        const linhaH = areaH / faixas.length;

        // Tudo deriva de linhaH (altura de cada faixa) — evita rótulos de faixas vizinhas
        // colidirem quando o gráfico ganha pouco espaço (várias seções na mesma página).
        // Antes usava tamanhos fixos (fonte 7.5, offsets fixos) que só cabiam quando havia
        // bastante altura disponível; com mais gráficos na grade, os textos se sobrepunham.
        const subH = Math.max(1.3, Math.min(4.2, linhaH * 0.24));
        const gapMeio = Math.max(0.4, Math.min(1.3, linhaH * 0.06));
        const fonteRotulo = Math.max(5.2, Math.min(7.5, linhaH * 0.5));
        const fonteValor = Math.max(5.3, Math.min(7.5, linhaH * 0.42));

        faixas.forEach((f, i) => {
            const base = topo + i * linhaH;
            const meio = base + linhaH / 2;
            doc.setFillColor(...(COR_SEVERIDADE[i] || COR.muted));
            doc.circle(x + 1.4, meio - 0.4, 1.1, 'F');
            doc.setFont('PublicSans', 'normal'); doc.setFontSize(fonteRotulo); doc.setTextColor(...COR.tintaSec);
            doc.text(doc.splitTextToSize(f.label, rotuloW - 6)[0], x + 4, meio + fonteRotulo * 0.15);

            // prioritários (acima do centro da faixa) — rótulo sempre centralizado na
            // própria barra, nunca a uma distância fixa de "meio".
            const wp = Math.max(0.5, (f.prioritarios / maxVal) * barMaxW);
            const yBarP = meio - gapMeio - subH;
            doc.setFillColor(...COR_PRIORITARIO);
            doc.roundedRect(barX, yBarP, wp, subH, 0.5, 0.5, 'F');
            doc.setFont('PublicSans', 'bold'); doc.setFontSize(fonteValor); doc.setTextColor(...COR.tinta);
            doc.text(String(f.prioritarios), barX + wp + 1.5, yBarP + subH / 2 + fonteValor * 0.15);

            // normais (abaixo do centro da faixa)
            const wn = Math.max(0.5, (f.normais / maxVal) * barMaxW);
            const yBarN = meio + gapMeio;
            doc.setFillColor(...COR.azul);
            doc.roundedRect(barX, yBarN, wn, subH, 0.5, 0.5, 'F');
            doc.text(String(f.normais), barX + wn + 1.5, yBarN + subH / 2 + fonteValor * 0.15);
        });
    }

    // link: quando informado (`{ label, pageNumber }`), desenha um link de navegação
    // centralizado no rodapé — usado pelo relatório conjunto para amarrar cada resumo à
    // sua tabela discriminada ("Ver tabela detalhada →") e cada tabela de volta ao seu
    // resumo ("← Voltar ao resumo"). Evita o glifo "↑" — fora da codificação padrão das
    // fontes do jsPDF e corrompe o texto renderizado.
    function desenharRodape(doc, titulo, quando, pw, ph, m, link) {
        doc.setFont('PublicSans', 'normal'); doc.setFontSize(8); doc.setTextColor(...COR.muted);
        doc.text(`${titulo}  •  Página ${doc.internal.getNumberOfPages()}`, m, ph - 6);
        doc.text(quando, pw - m, ph - 6, { align: 'right' });
        if (link) desenharLinkRodape(doc, link.label, link.pageNumber, pw, ph);
    }

    // Desenha só o link de navegação do rodapé (mesmo estilo de desenharRodape), sem
    // repetir o título/data — usado para "consertar" o link "Ver tabela detalhada →" de
    // um resumo já desenhado, depois que a tabela (mais à frente no PDF) é montada e sua
    // página passa a ser conhecida (ver PASSO 3 em gerarPDFConjunto).
    function desenharLinkRodape(doc, label, pageNumber, pw, ph) {
        doc.setFont('PublicSans', 'bold'); doc.setFontSize(8); doc.setTextColor(...COR.azul);
        const w = doc.getTextWidth(label);
        doc.textWithLink(label, pw / 2 - w / 2, ph - 6, { pageNumber });
    }

    // Distribui uma lista de gráficos numa grade de 2 colunas (span:2 ocupa a largura
    // toda) dentro da área (x, y, w, hDisponivel) informada — reutilizado tanto para os
    // gráficos da 1ª página do resumo quanto para os que foram para a 2ª página.
    function desenharGradeGraficos(doc, x, y, w, hDisponivel, charts) {
        if (!charts.length) return;
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
        const gap = 6;
        const colW = (w - gap) / 2;
        // Teto generoso (não 96mm como antes): com poucos gráficos numa página (ex.: só 2
        // depois de mover 3 para a 2ª página), a divisão por nLinhas já cresce sozinha —
        // um teto baixo só deixava espaço vazio sem necessidade.
        const chartH = Math.min(170, (hDisponivel - (nLinhas - 1) * 10) / nLinhas);
        charts.forEach(c => {
            const cx = x + c.pos.col * (colW + gap);
            const cy = y + c.pos.row * (chartH + 10);
            const cw = c.pos.span === 2 ? w : colW;
            if (c.tipo === 'faixas') desenharBarrasFaixas(doc, cx, cy, cw, chartH, c.titulo, c.faixas);
            else desenharBarras(doc, cx, cy, cw, chartH, c.titulo, c.itens, undefined, COR.aqua);
        });
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
            doc.setFont('PublicSans', 'bold'); doc.setFontSize(16); doc.setTextColor(...COR.tinta);
            doc.text(p.titulo, m, hy);
            hy += 8;
            if (contexto.competencia) {
                doc.setFont('PublicSans', 'bold'); doc.setFontSize(11.5); doc.setTextColor(...COR.azul);
                const linhas = doc.splitTextToSize('Competência: ' + contexto.competencia, uw);
                doc.text(linhas, m, hy);
                hy += linhas.length * 5.2 + 1.5;
            } else if (contexto.rotulo) {
                doc.setFont('PublicSans', 'bold'); doc.setFontSize(11.5); doc.setTextColor(...COR.azul);
                doc.text(contexto.rotulo, m, hy);
                hy += 7;
            }
            doc.setFont('PublicSans', 'normal'); doc.setFontSize(9); doc.setTextColor(...COR.tintaSec);
            let linhaInfo = `Extraído em ${hoje} às ${hora}  •  ${sub.length} registro(s)`;
            if (!contexto.competencia) {
                const fraseComp = fraseCompetencias(sub);
                if (fraseComp) linhaInfo += `  •  ${fraseComp}`;
            }
            // Quebra em várias linhas quando a lista de competências não cabe numa só —
            // sem isso o texto vazava a borda direita da página (visível sobretudo com 3+
            // competências combinadas no resumo geral).
            const linhasInfo = doc.splitTextToSize(linhaInfo, uw);
            doc.text(linhasInfo, m, hy);
            hy += (linhasInfo.length - 1) * 4.2 + 3;
            doc.setDrawColor(...COR.azul); doc.setLineWidth(0.5); doc.line(m, hy, pw - m, hy);

            // KPIs numéricos (média por dia só quando o relatório define mediaLabel)
            const kY = hy + 6;
            const prio = contarPrioritarios(sub);
            const kpis = [
                { titulo: p.atosTitulo, valor: String(sub.length), subs: [], acento: COR.azul },
            ];
            if (!p.semPrioridade) {
                kpis.push({ titulo: 'Prioritários pendentes', valor: String(prio), subs: [`${sub.length ? Math.round(prio / sub.length * 100) : 0}% do total`], acento: COR.vermelho });
            }
            if (p.mediaLabel) {
                // Quando há mais de uma competência, a média diária do "Resumo geral" é a
                // SOMA da média de cada competência (contexto.mediaSoma), não a média
                // calculada sobre os dados agrupados — tirar a média da média mistura as
                // distribuições de dias ativos de cada competência e distorce o resultado
                // (erro estatístico do tipo Simpson).
                const media = (typeof contexto.mediaSoma === 'number') ? contexto.mediaSoma : mediaPorDia(sub, p.dataCampo);
                kpis.push({ titulo: 'Média por dia', valor: media ? media.toLocaleString('pt-BR', { maximumFractionDigits: 1 }) : '—', subs: [p.mediaLabel], acento: COR.aqua });
            }
            const kW = (uw - (kpis.length - 1) * gap) / kpis.length;
            kpis.forEach((k, i) => desenharCard(doc, m + i * (kW + gap), kY, kW, 28, k.titulo, k.valor, k.subs, true, k.acento));

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

            // Gráficos: os marcados com pagina2 vão para uma segunda página do resumo,
            // ganhando a página inteira (útil para rankings maiores, ex.: 15 itens em vez
            // de 10, sem espremer os gráficos que ficam na página 1).
            const chartsTodos = [
                p.semPrioridade
                    ? { tipo: 'barras', span: 1, titulo: p.agingTitulo, itens: faixasPorPrioridade(sub, p.dataCampo, now).map(f => ({ label: f.label, valor: f.prioritarios + f.normais })), pagina2: false }
                    : { tipo: 'faixas', span: 1, titulo: p.agingTitulo, faixas: faixasPorPrioridade(sub, p.dataCampo, now), pagina2: false },
                // Gráficos de distribuição sem nenhum item qualificado (ex.: minValor, quando
                // nenhum processo tem mais de uma ocorrência) são omitidos inteiramente, em
                // vez de aparecer vazios.
                ...p.distribuicoes
                    .map(g => ({ tipo: 'barras', span: g.span || 1, titulo: g.titulo, itens: contarPorCampo(sub, g.campo, g.topN, g.limpar, g.semOutros, g.minValor), pagina2: !!g.pagina2 }))
                    .filter(c => c.itens.length),
            ];
            const chartsP1 = chartsTodos.filter(c => !c.pagina2);
            const chartsP2 = chartsTodos.filter(c => c.pagina2).map(c => ({ ...c, span: 2 })); // largura total na 2ª página

            const gY0 = aY + 28 + gap + 2;
            desenharGradeGraficos(doc, m, gY0, uw, ph - m - gY0, chartsP1);
            desenharRodape(doc, p.titulo, carimbo, pw, ph, m, comIndice);

            if (chartsP2.length) {
                doc.addPage();
                let hy2 = m + 2;
                doc.setFont('PublicSans', 'bold'); doc.setFontSize(16); doc.setTextColor(...COR.tinta);
                doc.text(p.titulo, m, hy2);
                hy2 += 8;
                if (contexto.competencia) {
                    doc.setFont('PublicSans', 'bold'); doc.setFontSize(11.5); doc.setTextColor(...COR.azul);
                    const linhas2 = doc.splitTextToSize('Competência: ' + contexto.competencia, uw);
                    doc.text(linhas2, m, hy2);
                    hy2 += linhas2.length * 5.2 + 1.5;
                } else if (contexto.rotulo) {
                    doc.setFont('PublicSans', 'bold'); doc.setFontSize(11.5); doc.setTextColor(...COR.azul);
                    doc.text(contexto.rotulo, m, hy2);
                    hy2 += 7;
                }
                doc.setFont('PublicSans', 'normal'); doc.setFontSize(9); doc.setTextColor(...COR.tintaSec);
                doc.text('Gráficos complementares', m, hy2);
                hy2 += 3;
                doc.setDrawColor(...COR.azul); doc.setLineWidth(0.5); doc.line(m, hy2, pw - m, hy2);

                const gY0b = hy2 + 6;
                desenharGradeGraficos(doc, m, gY0b, uw, ph - m - gY0b, chartsP2);
                desenharRodape(doc, p.titulo, carimbo, pw, ph, m, comIndice);
            }
        }

        // ═══ RESUMO POR COMPETÊNCIA (calculado antes do geral para permitir a soma das
        // médias — ver comentário no cálculo do KPI acima) ═══
        const porComp = new Map();
        dados.forEach(d => {
            const c = (d.competencia || '').trim();
            if (!c) return;
            if (!porComp.has(c)) porComp.set(c, []);
            porComp.get(c).push(d);
        });

        // ═══ RESUMO GERAL ═══
        const contextoGeral = { rotulo: 'Resumo geral' };
        if (porComp.size > 1) {
            contextoGeral.mediaSoma = [...porComp.values()].reduce((soma, sub) => soma + mediaPorDia(sub, p.dataCampo), 0);
        }
        desenharPaginaResumo(dados, contextoGeral, ehPrimeiraSecao);

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

        // ordenarPrioritarioPrimeiro (ex.: Conclusões): prioritárias primeiro; dentro de
        // cada grupo (e sempre que a flag não estiver ligada), da data mais antiga à mais
        // nova.
        const ordenados = dados.slice().sort((a, b) => {
            if (p.ordenarPrioritarioPrimeiro) {
                const pa = a.prioritario ? 0 : 1, pb = b.prioritario ? 0 : 1;
                if (pa !== pb) return pa - pb;
            }
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

        // Opções comuns ao autoTable, reaproveitadas tanto no caminho de tabela única
        // quanto no caminho agrupado (p.agruparPor) — mantém aparência/alinhamento
        // idênticos nos dois casos.
        const opcoesComuns = {
            columns: colunas.map((c, i) => ({ header: c.header, dataKey: 'k' + i })),
            margin: { left: m, right: m, top: m, bottom: 14 },
            theme: 'grid',
            styles: { font: 'PublicSans', fontSize: 7.5, cellPadding: 1.6, textColor: COR.tintaSec,
                      lineColor: COR.grade, lineWidth: 0.1, overflow: 'linebreak', valign: 'middle' },
            headStyles: { fillColor: COR.azul, textColor: [255, 255, 255], fontStyle: 'bold', fontSize: 8 },
            alternateRowStyles: { fillColor: COR.cartao },
            columnStyles,
        };
        const corpoDe = (itens) => itens.map(d => {
            const o = {};
            colunas.forEach((c, i) => { o['k' + i] = String(c.get(d, { now }) ?? ''); });
            return o;
        });

        if (!p.agruparPor) {
            // Comportamento padrão (idêntico ao de antes de existir p.agruparPor):
            // uma única tabela com todos os registros ordenados.
            doc.autoTable({
                ...opcoesComuns,
                body: corpoDe(ordenados),
                startY: tabInicioY,
                // Realça o número do processo dos prioritários
                didParseCell: (data) => {
                    if (data.section === 'body' && data.column.index === idxProcesso && ordenados[data.row.index] && ordenados[data.row.index].prioritario) {
                        data.cell.styles.textColor = COR_PRIORITARIO;
                        data.cell.styles.fontStyle = 'bold';
                    }
                },
                didDrawPage: () => desenharRodape(doc, p.titulo, carimbo, pw, ph, m, comIndice),
            });
        } else {
            // Tabela dividida em subtabelas por grupo (ex.: Apreensões por Tipo) — cada
            // grupo ganha um título de seção seguido de sua própria autoTable, na ordem
            // de p.ordemGrupos (fallback alfabético pt-BR pro que não estiver na lista).
            const grupos = new Map();
            ordenados.forEach(d => {
                const chave = String(d[p.agruparPor] || '').trim() || '(sem informação)';
                if (!grupos.has(chave)) grupos.set(chave, []);
                grupos.get(chave).push(d);
            });
            const ordem = p.ordemGrupos || [];
            const chaves = Array.from(grupos.keys()).sort((a, b) => {
                const ia = ordem.indexOf(a), ib = ordem.indexOf(b);
                if (ia !== -1 && ib !== -1) return ia - ib;
                if (ia !== -1) return -1;
                if (ib !== -1) return 1;
                return a.localeCompare(b, 'pt-BR');
            });

            let y = tabInicioY;
            chaves.forEach(chave => {
                const itens = grupos.get(chave);
                // Se o título da próxima seção não couber com folga pra ao menos uma
                // linha de tabela, começa página nova antes de desenhar o título.
                if (y > ph - m - 28) {
                    doc.addPage();
                    y = m + 8;
                }
                doc.setFont('PublicSans', 'bold'); doc.setFontSize(9.5); doc.setTextColor(...COR.azul);
                doc.text(`${chave} (${itens.length})`, m, y);

                doc.autoTable({
                    ...opcoesComuns,
                    body: corpoDe(itens),
                    startY: y + 4,
                    didParseCell: (data) => {
                        if (data.section === 'body' && data.column.index === idxProcesso && itens[data.row.index] && itens[data.row.index].prioritario) {
                            data.cell.styles.textColor = COR_PRIORITARIO;
                            data.cell.styles.fontStyle = 'bold';
                        }
                    },
                    didDrawPage: () => desenharRodape(doc, p.titulo, carimbo, pw, ph, m, comIndice),
                });
                y = doc.lastAutoTable.finalY + 8;
            });
        }

        return paginaInicial;
    }

    // Fonte Public Sans embutida em base64 (regular/negrito/itálico, subconjunto
    // latino) — usada em todos os PDFs no lugar da Helvetica padrão do jsPDF, para o
    // relatório não depender de conexão nem de fontes instaladas no computador de quem
    // gera o PDF. Registrada uma vez por documento em novoDocPDF().
    const FONTE_PUBLIC_SANS_REGULAR = 'AAEAAAAQAQAABAAAR0RFRizwKT0AAAJAAAAA6EdQT1OOkgtvAAAx6AAAMQZHU1VCs1+0hgAABsgAAASwT1MvMpEIZJgAAAHgAAAAYFNUQVTniMwXAAABmAAAAEhjbWFwsBMzOQAAEFgAAAZUZ2FzcAAAABAAAAEUAAAACGdseWarUdD/AABi8AAAeHpoZWFkIWZGzAAAAWAAAAA2aGhlYQ/2B6MAAAE8AAAAJGhtdHhU4um+AAAWrAAACbxsb2NhvVbcRQAAC3gAAATgbWF4cAJ/AMAAAAEcAAAAIG5hbWVaq3i1AAADKAAAA55wb3N0CeEYUAAAIGgAABF9cHJlcGgGjIUAAAEMAAAAB7gB/4WwBI0AAAEAAf//AA8AAQAAAm8AWgAHAGQABQABAAAAAAAAAAAAAAAAAAMAAQABAAAHbP4+AAAKMf7H/MQJwAABAAAAAAAAAAAAAAAAAAACbwABAAAAAgBC+X8vYl8PPPUAAwfQAAAAANs01VIAAAAA3pImAP7H/iAJwAhNAAAABgACAAAAAAAAAAEAAQAIAAIAAAAUAAIAAAAkAAJ3Z2h0AQAAAGl0YWwBEwABABQABAADAAEAAgEUAAAAAAABAAAAAwAAAAIBBAGQAAACvAAAAAQEeAGQAAUAAAUUBLAAAACWBRQEsAAAArwAjAJsAAAAAAAAAAAAAAAAoAAA/0AAIFsAAAAAAAAAAE5PTkUAwAAA+wIHbP4+AAAJCwHvIAABkwAAAAAECgWmAAAAIAADAAEAAgAuAAAADgAAANoACAACABgAEAABAAIBigGLAAEABAABArsAAQAEAAECgwACABwAAQAbAAEAHQA7AAEAPQBdAAEAXwCBAAEAhQCSAAEAlACuAAEAsAC0AAEAtgDlAAEA5wD+AAEBAAEXAAEBGQEcAAEBHgEjAAEBJQEpAAEBKwFHAAEBSwFYAAEBWgF1AAEBdwF7AAEBfQGJAAEBigGLAAIB9wH4AAECAQIBAAECBAIEAAECBgIHAAECNwI3AAECPgI/AAECQgJGAAMCSAJWAAMCZgJtAAMAAQABAAAACAABAAECUgAAAAwAlgADAAEECQAAAKoCXgADAAEECQABABYCSAADAAEECQACAA4COgADAAEECQADADoCAAADAAEECQAEACYB2gADAAEECQAFABoBwAADAAEECQAGACQBnAADAAEECQAOAXoAIgADAAEECQEAAAwAFgADAAEECQEEAA4COgADAAEECQETAAwACgADAAEECQEUAAoAAABSAG8AbQBhAG4ASQB0AGEAbABpAGMAVwBlAGkAZwBoAHQAUwBJAEwAIABPAHAAZQBuACAARgBvAG4AdAAgAEwAaQBjAGUAbgBzAGUALAAgAFYAZQByAHMAaQBvAG4AIAAxAC4AMQA6ACAAaAB0AHQAcABzADoALwAvAHMAYwByAGkAcAB0AHMALgBzAGkAbAAuAG8AcgBnAC8AYwBtAHMALwBzAGMAcgBpAHAAdABzAC8AcABhAGcAZQAuAHAAaABwAD8AcwBpAHQAZQBfAGkAZAA9AG4AcgBzAGkAJgBpAGQAPQBPAEYATABfAHcAZQBiADsAIABVAFMAVwBEAFMAIABNAG8AZABpAGYAaQBlAGQAIABWAGUAcgBzAGkAbwBuADoAIABoAHQAdABwAHMAOgAvAC8AZwBpAHQAaAB1AGIALgBjAG8AbQAvAHUAcwB3AGQAcwAvAHAAdQBiAGwAaQBjAC0AcwBhAG4AcwAvAGIAbABvAGIALwBtAGEAcwB0AGUAcgAvAEwASQBDAEUATgBTAEUALgBtAGQAUAB1AGIAbABpAGMAUwBhAG4AcwAtAFIAZQBnAHUAbABhAHIAVgBlAHIAcwBpAG8AbgAgADIALgAwADAAMQBQAHUAYgBsAGkAYwAgAFMAYQBuAHMAIABSAGUAZwB1AGwAYQByADIALgAwADAAMQA7AE4ATwBOAEUAOwBQAHUAYgBsAGkAYwBTAGEAbgBzAC0AUgBlAGcAdQBsAGEAcgBSAGUAZwB1AGwAYQByAFAAdQBiAGwAaQBjACAAUwBhAG4AcwBDAG8AcAB5AHIAaQBnAGgAdAAgADIAMAAxADUAIABUAGgAZQAgAFAAdQBiAGwAaQBjACAAUwBhAG4AcwAgAFAAcgBvAGoAZQBjAHQAIABBAHUAdABoAG8AcgBzACAAKABoAHQAdABwAHMAOgAvAC8AZwBpAHQAaAB1AGIALgBjAG8AbQAvAHUAcwB3AGQAcwAvAHAAdQBiAGwAaQBjAC0AcwBhAG4AcwApAAAAAQAAAAoAxAFiAAJERkxUAKBsYXRuAA4AfAAIQVpFIACWQ0FUIABkQ1JUIACWS0FaIACWTU9MIABMUk9NIACW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ALIBeQC4AX8AuQC/AYYAwQGIAMABhwB0AToAowFqAH0BQwAYANoAGwDdAH8BRQAPANEAFADWADIA9AA4APoASgEOAFABFABvATUAewFBAIkBTwCLAVEAngFlAKoBcQCRAVcAmAFeAkECQAJeAlgCXwJjAmACWwJEAkUCSAJMAk0CSgJDAkICTgJLAkYCSQAnAOkARQEIAGMBKQCKAVAAkgFYAJkBXwC0AXsAsQF4ALMBegDCAYkAEQDTABMA1QAKAMwADADOAA0AzwAOANAACwDNAAQAxgAGAMgABwDJAAgAygAFAMcANQD3ADcA+QA7AP0ALQDvAC8A8QAwAPIAMQDzAC4A8ABPARMATQERAHEBNwBzATkAagEwAGwBMgBtATMAbgE0AGsBMQB1ATsAdwE9AHgBPgB5AT8AdgE8AKABZwCiAWkApAFrAKYBbQCnAW4AqAFvAKUBbAC7AYIAugGBALwBgwC9AYQB7AHtAegB6gHrAekCPAI9AdUCAgIAAj4CNgIjAicCKQIVAhICKgIdAhwFXAC/BYwAUwWMAFMFjABTBYwAUwWMAFMFjABTBYwAUwWMAFMFjABTBYwAUwWMAFMFjABTBYwAUwWMAFMFjABTBYwAUwWMAFMFjABTBYwAUwWMAFMFjABTBYwAUwWMAFMFjABTBYwAUwedAAAHnQAABT0AvwVhAGwFYQBsBWEAbAVhAGwFYQBsBWEAbAVrAL8FhgBGBWsAvwWGAEYFawC/BNkAvwTZAL8E2QC/BNkAvwTZAL8E2QC/BNkAvwTZAL8E2QC/BNkAvwTZAJIE2QC/BNkAvwTZAL8E2QC/BNkAvwTZAL8E2QC/BNkAvwTZAL8EpAC/BboAbAW6AGwFugBsBboAbAW6AGwF/AC/Bi8AGwX8AL8F/AC/AjYAvwI2AL8CNv/tAjb/4AI2/yUCNv//AjYAvgI2AL8CNv/IAjYAkwI2/+kCNv/xAjYAJAI2//IDCQBTAwkAUwVUAL8FVAC/BIMAvwSDAL8EgwC/BIMAvwSDAL8EigAUBusAvwXkAL8F5AC/BeQAvwXkAL8F5AC/BeQAvwXkAL8FvABsBbwAbAW8AGwFvABsBbwAbAW8AGwFvABsBbwAbAW8AGwFvABsBbwAbAW8AGwFvABsBbwAbAYjAGwGIwBsBiMAbAYjAGwGIwBsBiMAbAW8AGwFvABsBbwAbAW8AGwF3QA4Bd0AOAW8AGwIvABsBPkAvwUVAL8F1wBsBUkAvwVJAL8FSQC/BUkAvwVJAJsFSQC/BUkAvwU8AHQFPAB0BTwAdAU8AHQFPAB0BTwAdAU8AHQFugBsBKMALgSjAC4EowAuBKMALgSjAC4EowAuBXkAvwV5AL8FeQC/BXkAvwV5AL8FeQC/BXkAvwV5AL8FeQC/BlQAvwZUAL8GVAC/BlQAvwZUAL8GUAC/BXkAvwV5AL8FeQC/BXkAvwV5AL8FeQC/BUEALgdzAC4HcwAuB3MALgdzAC4HcwAuBWAAUwS+ABgEvgAYBL4AGAS+ABgEvgAYBL4AGAS+ABgEvgAYBPoAfwT6AH8E+gB/BPoAfwT6AH8EUgBmBFIAZgRSAGYEUgBmBFIAZgRSAGYEUgBmBFIAZgRSAGYEUgBmBFIAZgRSAGYEUgBmBFIAZgRSAEAEUgBmBFIAZgRSAGYEUgBmBFIAZgRSAGYEUgBmBFIAZgRSAGYEUgBmByQAZgckAGYEqQCZBD8AXQQ/AF0EPwBdBD8AXQQ/AF0EPwBdBKkAXQRrAF0EnABdBNIAXQSpAF0EbQBmBG0AZgRtAGYEZgBmBG0AZgRtAGYEbQBmBG0AZgRtAGYEbQBmBG0ASgRtAGYEbQBmBG0AZgRtAGYEbQBmBG0AZgRtAGYEbQBmBG0AZgRtAGcDGABIBN0ASwTdAEsE3QBLBN0ASwTdAEsEgQCZBJkAKASB/7YEgQCZAe8AmQHkAJkB5ACZAeT/xAHk/7cB5P78AlsADwHkAJUB7wCZAeT/5gHkAGoCNf/nAjn/8QHk//cCL//rAg3/0gIJ/9ICCf/SBGMAmQRjAJkEeACZAl4AmQJeAJkCXgCZAl4AmAK9AJkCwQBGBsIAmQSCAJkEggCZBIIAmQSCAJkEggCZBJMAmQSCAJkEawBdBGsAXQRrAF0EawBdBGsAXQRrAF0EawBdBGsAXQRrAF0EawA/BGsAXQRrAF0EawBdBGsAXQSzAF0EswBdBLMAXQSzAF0EswBdBLMAXQRrAF0EawBdBGsAXQRrAF0ElAAoBJQAKARrAF0HaABdBKcAmQSbAJkEpwBdAw4AmQMOAJkDDgBpAw4AhQMO/68DDgCWAw4AcwP0AF8D9ABfA/QAXwP0AF8D9ABfA/QAXwP0AF8EmwCZAwoAXAMKAFwDTQBcAwoAXAMKAFwDCgBcBIkAiASJAIgEiQCIBIkAiASJAIgEiQBIBIkAiASJAIgEiQCIBIkAiAUYAIgFGACIBRgAiAUYAIgFGACIBRgAiASJAIgEiQCIBIkAiASJAIgEiQCIBIkAiAPxAC8GGAA0BhgANAYYADQGGAA0BhgANARpAD8ELgAvBC4ALwQuAC8ELgAvBC4ALwQuAC8ELgAvBC4ALwPBAEwDwQBMA8EATAPBAEwDwQBMBQcASAV2AEgDMgBmAywAXQUmADUEyABgAy4AXQSoAF0E7ABgBOkAXQUBAGAE3QBgBLEAXQUPAGAE4QBjBXgAuAV4AMoFeADVBXgArwV4AJMFeAC0BXgArwV4AJQFeACVBXgAtQL3ADcB3wAuAt4ALgMjAEEDEQAuAxwANwMFADcCyAAuAz0AQQMIADkC9wA3Ad8ALgLeAC4DIwBBAxEALgMcADcDBQA3AsgALgM9AEEDCAA5AvcANwHfAC4C3gAuAyMAQQMRAC4DHAA3AwUANwLIAC4DPQBBAwgAOQHQ/scGxQAuBtwALggNAC4GLwAuBy0AQQb2AC4H8wBBB+wANwc8AC4BwgBuAeMAVwHDAG4B4ABuBf8AbgH2AIgB9wCIBBIALgQMAD8BywBuAnkAdAMUADUEtgBJAr0AGwK9ABsCWAC+AlgArAJdAIsCagAnAo8AEQKYACYCXgCzAl4ANgLzAHwC8wB8BFYAfAi6AXMECwB7BdQAAAHjAFcDUgBXA0YAagM1AGQB3wBqAc8AZAONAFEDvQB7AiwAUQIrAGsDUwCIAfAAiAHmAAAB5gAAAdMAAAQ9AGsFYQBsBDcAZwUlAHcD+QBTBK//7QOQ/3cEwP/yBMoAHgRvAB4F8AANBJEAFQTKAB4ICQAbBMcACQV4AQEFeABtBXgAoQV4APoFeAB2BXgArgV4AIUFeACWBXgArgV4AJIFeACWBXgAYgJ8AL4EpgBgBFUAewRVAHsEOAB7BFUAewRVAHsEVQB7BEMAewRDAHsEVQB7BFUAewRVAHsEdgB7BHAAdQRVAHsDsgBZBIYAagW6AGEDcAARBfsAdAV/ACUFIAAhBCEANQTxAB4EnABgBJMAiAcQAGoKMQBqA9EAXgZ+AF0FiQBXBF4AXQPVAEkGYgBGBDYARgZiAEYHIQAYA54AkgIOAL8CDgC/A/MALwOTAB8DkwAf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cmN1bWZsZXgHdW5pMDEyMgpHZG90YWNjZW50BEhiYXILSGNpcmN1bWZsZXgHdW5pMUUyNAZJYnJldmUHdW5pMDIwOAd1bmkxRUNBB3VuaTFFQzgHdW5pMDIwQQdJbWFjcm9uB0lvZ29uZWsGSXRpbGRlC0pjaXJjdW1mbGV4B3VuaTAxMzYGTGFjdXRlBkxjYXJvbgd1bmkwMTNCBExkb3QGTmFjdXRlBk5jYXJvbgd1bmkwMTQ1B3VuaTFFNDQDRW5nBk9icmV2ZQd1bmkxRUQwB3VuaTFFRDgHdW5pMUVEMgd1bmkxRUQ0B3VuaTFFRDYHdW5pMDIwQwd1bmkxRUNDB3VuaTFFQ0UFT2hvcm4HdW5pMUVEQQd1bmkxRUUyB3VuaTFFREMHdW5pMUVERQd1bmkxRUUwDU9odW5nYXJ1bWxhdXQHdW5pMDIwRQdPbWFjcm9uB3VuaTAxRUELT3NsYXNoYWN1dGUGUmFjdXRlBlJjYXJvbgd1bmkwMTU2B3VuaTAyMTAHdW5pMUU1QQd1bmkwMjEyBlNhY3V0ZQtTY2lyY3VtZmxleAd1bmkwMjE4B3VuaTFFNjIHdW5pMDE4RgRUYmFyBlRjYXJvbgd1bmkwMTYyB3VuaTAyMUEHdW5pMUU2QwZVYnJldmUHdW5pMDIxNAd1bmkxRUU0B3VuaTFFRTYFVWhvcm4HdW5pMUVFOAd1bmkxRUYwB3VuaTFFRUEHdW5pMUVFQwd1bmkxRUVFDVVodW5nYXJ1bWxhdXQHdW5pMDIxNgdVbWFjcm9uB1VvZ29uZWsFVXJpbmcGVXRpbGRlBldhY3V0ZQtXY2lyY3VtZmxleAlXZGllcmVzaXMGV2dyYXZlC1ljaXJjdW1mbGV4B3VuaTFFRjQGWWdyYXZlB3VuaTFFRjYHdW5pMUVGOAZaYWN1dGUKWmRvdGFjY2VudAd1bmkxRTkyBmFicmV2ZQd1bmkxRUFGB3VuaTFFQjcHdW5pMUVCMQd1bmkxRUIzB3VuaTFFQjUHdW5pMUVBNQd1bmkxRUFEB3VuaTFFQTcHdW5pMUVBOQd1bmkxRUFCB3VuaTAyMDEHdW5pMUVBMQd1bmkxRUEzB3VuaTAyMDMHYW1hY3Jvbgdhb2dvbmVrCmFyaW5nYWN1dGUHYWVhY3V0ZQtjY2lyY3VtZmxleApjZG90YWNjZW50BmRjYXJvbgd1bmkxRTBEBmVicmV2ZQZlY2Fyb24HdW5pMUVCRgd1bmkxRUM3B3VuaTFFQzEHdW5pMUVDMwd1bmkxRUM1B3VuaTAyMDUKZWRvdGFjY2VudAd1bmkxRUI5B3VuaTFFQkIHdW5pMDIwNwdlbWFjcm9uB2VvZ29uZWsHdW5pMUVCRAd1bmkwMjU5C2djaXJjdW1mbGV4B3VuaTAxMjMKZ2RvdGFjY2VudARoYmFyC2hjaXJjdW1mbGV4B3VuaTFFMjUGaWJyZXZlB3VuaTAyMDkJaS5sb2NsVFJLB3VuaTFFQ0IHdW5pMUVDOQd1bmkwMjBCB2ltYWNyb24HaW9nb25lawZpdGlsZGUHdW5pMDIzNwtqY2lyY3VtZmxleAd1bmkwMTM3DGtncmVlbmxhbmRpYwZsYWN1dGUGbGNhcm9uB3VuaTAxM0MEbGRvdAZuYWN1dGUGbmNhcm9uB3VuaTAxNDYHdW5pMUU0NQNlbmcGb2JyZXZlB3VuaTFFRDEHdW5pMUVEOQd1bmkxRUQzB3VuaTFFRDUHdW5pMUVENwd1bmkwMjBEB3VuaTFFQ0QHdW5pMUVDRgVvaG9ybgd1bmkxRURCB3VuaTFFRTMHdW5pMUVERAd1bmkxRURGB3VuaTFFRTENb2h1bmdhcnVtbGF1dAd1bmkwMjBGB29tYWNyb24HdW5pMDFFQgtvc2xhc2hhY3V0ZQZyYWN1dGUGcmNhcm9uB3VuaTAxNTcHdW5pMDIxMQd1bmkxRTVCB3VuaTAyMTMGc2FjdXRlC3NjaXJjdW1mbGV4B3VuaTAyMTkHdW5pMUU2MwR0YmFyBnRjYXJvbgd1bmkwMTYzB3VuaTAyMUIHdW5pMUU2RAZ1YnJldmUHdW5pMDFENAd1bmkwMjE1B3VuaTFFRTUHdW5pMUVFNwV1aG9ybgd1bmkxRUU5B3VuaTFFRjEHdW5pMUVFQgd1bmkxRUVEB3VuaTFFRUYNdWh1bmdhcnVtbGF1dAd1bmkwMjE3B3VtYWNyb24HdW9nb25lawV1cmluZwZ1dGlsZGUGd2FjdXRlC3djaXJjdW1mbGV4CXdkaWVyZXNpcwZ3Z3JhdmULeWNpcmN1bWZsZXgHdW5pMUVGNQZ5Z3JhdmUHdW5pMUVGNwd1bmkxRUY5BnphY3V0ZQp6ZG90YWNjZW50B3VuaTFFOTMHemVyby50ZgZvbmUudGYGdHdvLnRmCHRocmVlLnRmB2ZvdXIudGYHZml2ZS50ZgZzaXgudGYIc2V2ZW4udGYIZWlnaHQudGYHbmluZS50Zgl6ZXJvLmRub20Ib25lLmRub20IdHdvLmRub20KdGhyZWUuZG5vbQlmb3VyLmRub20JZml2ZS5kbm9tCHNpeC5kbm9tCnNldmVuLmRub20KZWlnaHQuZG5vbQluaW5lLmRub20JemVyby5udW1yCG9uZS5udW1yCHR3by5udW1yCnRocmVlLm51bXIJZm91ci5udW1yCWZpdmUubnVtcghzaXgubnVtcgpzZXZlbi5udW1yCmVpZ2h0Lm51bXIJbmluZS5udW1yB3VuaTIwNzAHdW5pMDBCOQd1bmkwMEIyB3VuaTAwQjMHdW5pMjA3NAd1bmkyMDc1B3VuaTIwNzYHdW5pMjA3Nwd1bmkyMDc4B3VuaTIwNzkHdW5pMjE1Mwd1bmkyMTU0CW9uZWVpZ2h0aAx0aHJlZWVpZ2h0aHMLZml2ZWVpZ2h0aHMMc2V2ZW5laWdodGhzCXBlcmlvZC50Zghjb21tYS50Zgd1bmkwMEFECmVtZGFzaC5hbHQHdW5pMDBBMAJDUg1jb2xvbm1vbmV0YXJ5BGRvbmcERXVybwRsaXJhB3VuaTIwQkEHdW5pMjBBNgd1bmkyMEI5B3VuaTIwQTkHY2VudC50ZhBjb2xvbm1vbmV0YXJ5LnRmCWRvbGxhci50Zgdkb25nLnRmB0V1cm8udGYJZmxvcmluLnRmCGZyYW5jLnRmB2xpcmEudGYKdW5pMjBCQS50Zgp1bmkyMEI5LnRmC3N0ZXJsaW5nLnRmBnllbi50Zgd1bmkyMjE5B3VuaTIyMTUIZW1wdHlzZXQHdW5pMjEyNgd1bmkyMjA2B3VuaTAwQjUHdW5pMjExNwd1bmkyMTEzB3VuaTIxMTYHdW5pMjEyMAd1bmkwMkJDB3VuaTAyQkIHdW5pMDMwOAd1bmkwMzA3CWdyYXZlY29tYglhY3V0ZWNvbWIHdW5pMDMwQgt1bmkwMzBDLmFsdAd1bmkwMzAyB3VuaTAzMEMHdW5pMDMwNgd1bmkwMzBBCXRpbGRlY29tYgd1bmkwMzA0DWhvb2thYm92ZWNvbWIHdW5pMDMwRgd1bmkwMzExB3VuaTAzMTIHdW5pMDMxQgxkb3RiZWxvd2NvbWIHdW5pMDMyNgd1bmkwMzI3B3VuaTAzMjgQZ3JhdmVjb21iLm5hcnJvdxBhY3V0ZWNvbWIubmFycm93C3VuaTAzMDYwMzAxC3VuaTAzMDYwMzAwC3VuaTAzMDYwMzA5C3VuaTAzMDYwMzAzC3VuaTAzMDIwMzAxC3VuaTAzMDIwMzAwC3VuaTAzMDIwMzA5C3VuaTAzMDIwMzAzB3VuaTAwMDAAAAAAAQAAAAoAKABQAAJERkxUAA5sYXRuAA4ABAAAAAD//wADAAAAAQACAANrZXJuACJtYXJrABpta21rABQAAAABAAMAAAACAAEAAgAAAAEAAAAEETAQpAAkAAoABgAQAAEACgAAAAERFhEWAAERCgAMAAERBAAEAAAAAQAIAAEQYg7YAAQPbAAMAX0OxgAADsAOug7GAAAOwA66DsYAAA7ADroOxgAADsAOug7GAAAOwA66DsYAAA7ADroOxgAADsAOug7GAAAOwA66DsYAAA7ADroOxgAADsAOug7GAAAOwA66DsYAAA7ADroOxgAADsAOug7GAAAOwA66DsYAAA7ADroOxgAADsAOug7GAAAOwA66DsYAAA7ADroOxgAADsAOug7GAAAOwA66DsYAAA7ADroOxgAADsAOug7GAAAOwA66DsYAAA7ADroOxgAADsAOug60AAAAAAAADrQAAAAAAAAOrgAADqgAAA6uAAAOqAAADq4AAA6oAAAOrgAADqgAAA6uAAAOqAAADq4AAA6oAAAOogAADsAAAA6cAAAOlgAADqIAAA7AAAAOnAAADpYAAA6iAAAOwAAADpAAAA6KDoQOkAAADooOhA6QAAAOig6EDpAAAA6KDoQOkAAADooOhA6QAAAOig6EDpAAAA6KDoQOkAAADooOhA6QAAAOig6EDpAAAA6KDoQOkAAADooOhA6QAAAOig6EDpAAAA6KDoQOkAAADooOhA6QAAAOig6EDpAAAA6KDoQOkAAADooOhA6QAAAOig6EDpAAAA6KDoQOkAAADooOhA5+AAAOeAAADn4AAA54AAAOfgAADngAAA5+AAAOeAAADn4AAA54AAAOcgAADmwAAA5mAAAOYAAADnIAAA5sAAAOcgAADmwAAA5aAAAOVA5ODloAAA5UDk4OWgAADlQOTg5aAAAOVA5ODloAAA5UDk4OWgAADlQOTg5aAAAOVA5ODloAAA5UDk4OWgAADlQOTg5aAAAOVA5ODloAAA5UDk4OWgAADlQOTg5aAAAOVA5ODloAAA5UDk4OSAAAAAAAAA5IAAAAAAAADkIAAA48AAAOQgAADjwAAA42AAAOMAAADjYAAA4wAAAONgAADjAAAA42AAAOMAAADjYAAA4wAAAOKgAADiQAAA4eAAAOeAAADh4AAA54AAAOHgAADngAAA4eAAAOeAAADh4AAA54AAAOHgAADngAAA4eAAAOeAAADhgOEg6WDgwOGA4SDpYODA4YDhIOlg4MDhgOEg6WDgwOGA4SDpYODA4YDhIOlg4MDhgOEg6WDgwOGA4SDpYODA4YDhIOlg4MDhgOEg6WDgwOGA4SDpYODA4YDhIOlg4MDhgOEg6WDgwOGA4SDpYODA4YDhIOlg4MDhgOEg6WDgwOGA4SDpYODA4YDhIOlg4MDhgOEg6WDgwOGA4SDpYODA4YDhIOlg4MDhgOEg6WDgwOGA4SDpYODA4YDhIOlg4MDgYAAAAAAAAOBgAAAAAAAA4YDhIOlg4MDgAAAAAAAAAN+gAADfQAAA36AAAN9AAADfoAAA30AAAN+gAADfQAAA36AAAN9AAADfoAAA30AAAN+gAADfQAAA3uAAAN6AAADe4AAA3oAAAN7gAADegAAA3uAAAN6AAADe4AAA3oAAAN7gAADegAAA3uAAAN6AAADeIAAA3cAAAN4gAADdwAAA3iAAAN3AAADeIAAA3cAAAN4gAADdwAAA3iAAAN3AAADdYN0A3KDcQN1g3QDcoNxA3WDdANyg3EDdYN0A3KDcQN1g3QDcoNxA3WDdANyg3EDdYN0A3KDcQN1g3QDcoNxA3WDdANyg3EDdYN0A3KDcQN1g3QDcoNxA3WDdANyg3EDdYN0A3KDcQN1g3QDcoNxA3WDdANyg3EDdYN0A3KDcQN1g3QDcoNxA3WDdANyg3EDdYN0A3KDcQN1g3QDcoNxA3WDdANyg3EDb4AAAAAAAANvgAAAAAAAA2+AAAAAAAADb4AAAAAAAANvgAAAAAAAA24AAANsgAADbgAAA2yAAANuAAADbIAAA24AAANsgAADbgAAA2yAAANuAAADbIAAA24AAANsgAADbgAAA2yAAANrAAADaYAAA2sAAANpgAADawAAA2mAAANrAAADaYAAA2sAAANpgAADaAAAA2aDZQNoAAADZoNlA2gAAANmg2UDaAAAA2aDZQNoAAADZoNlA2gAAANmg2UDaAAAA2aDZQNoAAADZoNlA2gAAANmg2UDaAAAA2aDZQNoAAADZoNlA2gAAANmg2UDaAAAA2aDZQNoAAADZoNlA2gAAANmg2UDaAAAA2aDZQNoAAADZoNlA2gAAANmg2UDaAAAA2aDZQNoAAADZoNlA2gAAANmg2UDaAAAA2aDZQNoAAADZoNlA2gAAANmg2UDaAAAA2aDZQNjgAADZoNlA2OAAANmg2UAAAAAA2IAAANggAADXwAAA2CAAANfAAADYIAAA18AAANggAADXwAAA2CAAANfAAADYIAAA18AAAAAAAADXYAAAAAAAANdgAAAAAAAA12AAAAAAAADXYAAA1wAAANag1kDXAAAA1qDWQNcAAADWoNZA1wAAANag1kDXAAAA1qDWQNcAAADWoNZA1wAAANag1kDXAAAA1qDWQNcAAADWoNZA1wAAANag1kDXAAAA1qDWQNcAAADWoNZA1wAAANag1kDXAAAA1qDWQNcAAADWoNZA1wAAANag1kDXAAAA1qDWQNcAAADWoNZA1wAAANag1kDXAAAA1qDWQNXgAADVgNUg1MAAAAAAAADUwAAAAAAAANTAAAAAAAAA1MAAAAAAAADUwAAAAAAAANRgAADUAAAA06AAANNAAADUYAAA1AAAANRgAADUAAAAAAAAANLgAADSgAAAAADSINKAAAAAANIg0oAAAAAA0iDSgAAAAADSINKAAAAAANIg0cAAAAAA0WDSgAAAAADSIAAAAADS4AAA0oAAAAAA0iDSgAAAAADSINEAAAAAANCg0EAAAAAAz+DSgAAAAADSIM+AAAAAAM8gzsAAAAAAAADOwAAAAAAAAAAAAADOYAAAAAAAAM5gAADOAAAAzaAAAM4AAADNoAAAzgAAAM2gAADOAAAAzaAAAM4AAADNoAAAzUAAAMzgAADMgAAAzCAAAMyAAADMIAAAzIAAAMwgAADMgAAAzCAAAMyAAADMIAAAzIAAAMwgAADLwMtgywDKoMvAy2DLAMqgy8DLYMsAyqDLwMtgywDKoMvAy2DLAMqgy8DLYMsAyqDLwMtgywDKoMvAy2DLAMqgy8DLYMsAyqDLwMtgywDKoMvAy2DLAMqgy8DLYMsAyqDLwMtgywDKoMvAy2DLAMqgy8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ACQAAP//AL/+mQT/BaYCJgAjAAAABwJTAmAAAAABAL8AAARiBaYACwAAMxEhByERIRUhEQUVvwOQAf0rAo39cgLqBaaj/iii/h4Cpf//AL8AAARiB0ICJgAoAAAABwJFAi8BnP//AL8AAARiBvUCJgAoAAAABwJKAUYBmf//AL8AAARiBwACJgAoAAAABwJJAUQBnP//AL8AAARiBwACJgAoAAAABwJIAUQBnP//AL8AAASAB+4CJgAoAAAABwJqAUQBnP//AL/+mQRiBwACJgAoAAAAJwJTAkYAAAAHAkgBRAGc//8AvwAABGIH7wImACgAAAAHAmsBRAGc//8AvwAABGIINwImACgAAAAHAmwBSQGc//8AvwAABGIIJAImACgAAAAHAm0BRgGc//8AkgAABGIHQgImACgAAAAHAk8AiAGc//8AvwAABGIHQgImACgAAAAHAkIBYgGc//8AvwAABGIHQgImACgAAAAHAkMCIgGc//8Av/6ZBGIFpgImACgAAAAHAlMCRgAA//8AvwAABGIHQgImACgAAAAHAkQBKwGc//8AvwAABGIHhQImACgAAAAHAk4ALQGL//8AvwAABGIHLAImACgAAAAHAlABTgGc//8AvwAABGIG8QImACgAAAAHAk0BRQGc//8Av/6xBJ0FpgImACgAAAAHAlYDA//9//8AvwAABGIHGAImACgAAAAHAkwBNwGcAAEAvwAABEQFpgAJAAAzESEHIREhFSERvwOFAf06An79ggWmov4ioP16AAEAbP/sBREFugAqAAAFIiQCNTQSJDMyHgIXIy4CIyIGAhUUHgIzMj4CNzchNQURIzUOAgLSvv7slJcBGL9twJRZBsYOYJNcfb1pQHObWluHWjECBv6gAiCLJnSnFK8BSursAU6xTIOoW1CKVHv/AMia3IlBQGV1NoCIAf0W3D5uRP//AGz/7AURBvUCJgA9AAAABwJKAZMBmf//AGz/7AURBwACJgA9AAAABwJIAZEBnP//AGz+hAURBboCJgA9AAAABwJUAmQAAP//AGz/7AURB0ICJgA9AAAABwJDAm4BnAABAL8AAAU9BaYACwAAMxEzESERMxEjESERv7kDDLm5/PQFpv2YAmj6WgKb/WUAAgAbAAAGFAWmAAMADwAAEzUhFQERMxEhETMRIxEhERsF+frLuQMMubn89ARY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';
    const FONTE_PUBLIC_SANS_BOLD = 'AAEAAAAQAQAABAAAR0RFRi0gKT0AAAI8AAAA6EdQT1OOCh1eAAAxzAAAMRJHU1VCs1+0hgAABqwAAASwT1MvMpI0ZKgAAAHcAAAAYFNUQVTl9MwaAAABmAAAAERjbWFwsBMzOQAAEDwAAAZUZ2FzcAAAABAAAAEUAAAACGdseWbWB3A/AABi4AAAeUxoZWFkIblHGAAAAWAAAAA2aGhlYRBIB2YAAAE8AAAAJGhtdHjFhsK5AAAWkAAACbxsb2Nh7kgNZwAAC1wAAATgbWF4cAJ/AMAAAAEcAAAAIG5hbWVX7HWpAAADJAAAA4Zwb3N0CeEYUAAAIEwAABF9cHJlcGgGjIUAAAEMAAAAB7gB/4WwBI0AAAEAAf//AA8AAQAAAm8AWgAHAGQABQABAAAAAAAAAAAAAAAAAAMAAQABAAAHbP4+AAAKkf6y/CcKJwABAAAAAAAAAAAAAAAAAAACbwABAAAAAgBCZp2+oF8PPPUAAwfQAAAAANs01VIAAAAA3pImAP6y/iAKJwiZAAEABgACAAAAAAAAAAEAAQAIAAIAAAAUAAIAAAAkAAJ3Z2h0AQAAAGl0YWwBEwABABQABAADAAEAAgEUAAAAAAABAAAAAQAAAAABBwK8AAAABASoArwABQAABRQEsAAAAJYFFASwAAACvACMAmwAAAAAAAAAAAAAAACgAAD/QAAgWwAAAAAAAAAATk9ORQCgAAD7Agds/j4AAAkLAe8gAAGTAAAAAAQKBaYAAAAgAAMAAQACAC4AAAAOAAAA2gAIAAIAGAAQAAEAAgGKAYsAAQAEAAEC1gABAAQAAQKYAAIAHAABABsAAQAdADsAAQA9AF0AAQBfAIEAAQCFAJIAAQCUAK4AAQCwALQAAQC2AOUAAQDnAP4AAQEAARcAAQEZARwAAQEeASMAAQElASkAAQErAUcAAQFLAVgAAQFaAXUAAQF3AXsAAQF9AYkAAQGKAYsAAgH3AfgAAQIBAgEAAQIEAgQAAQIGAgcAAQI3AjcAAQI+Aj8AAQJCAkYAAwJIAlYAAwJmAm0AAwABAAEAAAAIAAEAAQJSAAAADACWAAMAAQQJAAAAqgJGAAMAAQQJAAEAFgIwAAMAAQQJAAIACAIoAAMAAQQJAAMANAH0AAMAAQQJAAQAIAHUAAMAAQQJAAUAGgG6AAMAAQQJAAYAHgGcAAMAAQQJAA4BegAiAAMAAQQJAQAADAAWAAMAAQQJAQcACAIoAAMAAQQJARMADAAKAAMAAQQJARQACgAAAFIAbwBtAGEAbgBJAHQAYQBsAGkAYwBXAGUAaQBnAGgAdABTAEkATAAgAE8AcABlAG4AIABGAG8AbgB0ACAATABpAGMAZQBuAHMAZQAsACAAVgBlAHIAcwBpAG8AbgAgADEALgAxADoAIABoAHQAdABwAHMAOgAvAC8AcwBjAHIAaQBwAHQAcwAuAHMAaQBsAC4AbwByAGcALwBjAG0AcwAvAHMAYwByAGkAcAB0AHMALwBwAGEAZwBlAC4AcABoAHAAPwBzAGkAdABlAF8AaQBkAD0AbgByAHMAaQAmAGkAZAA9AE8ARgBMAF8AdwBlAGIAOwAgAFUAUwBXAEQAUwAgAE0AbwBkAGkAZgBpAGUAZAAgAFYAZQByAHMAaQBvAG4AOgAgAGgAdAB0AHAAcwA6AC8ALwBnAGkAdABoAHUAYgAuAGMAbwBtAC8AdQBzAHcAZABzAC8AcAB1AGIAbABpAGMALQBzAGEAbgBzAC8AYgBsAG8AYgAvAG0AYQBzAHQAZQByAC8ATABJAEMARQBOAFMARQAuAG0AZABQAHUAYgBsAGkAYwBTAGEAbgBzAC0AQgBvAGwAZABWAGUAcgBzAGkAbwBuACAAMgAuADAAMAAxAFAAdQBiAGwAaQBjACAAUwBhAG4AcwAgAEIAbwBsAGQAMgAuADAAMAAxADsATgBPAE4ARQA7AFAAdQBiAGwAaQBjAFMAYQBuAHMALQBCAG8AbABkAEIAbwBsAGQAUAB1AGIAbABpAGMAIABTAGEAbgBzAEMAbwBwAHkAcgBpAGcAaAB0ACAAMgAwADEANQAgAFQAaABlACAAUAB1AGIAbABpAGMAIABTAGEAbgBzACAAUAByAG8AagBlAGMAdAAgAEEAdQB0AGgAbwByAHMAIAAoAGgAdAB0AHAAcwA6AC8ALwBnAGkAdABoAHUAYgAuAGMAbwBtAC8AdQBzAHcAZABzAC8AcAB1AGIAbABpAGMALQBzAGEAbgBzACkAAAABAAAACgDEAWIAAkRGTFQAoGxhdG4ADgB8AAhBWkUgAJZDQVQgAGRDUlQgAJZLQVogAJZNT0wgAExST00gAJZUQVQgAJZUUksgADQAAP//AAkAAAABAAMABAAF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AWoAfQFDABgA2gAbAN0AfwFFAA8A0QAUANYAMgD0ADgA+gBKAQ4AUAEUAG8BNQB7AUEAiQFPAIsBUQCeAWUAqgFxAJEBVwCYAV4CQQJAAl4CWAJfAmMCYAJbAkQCRQJIAkwCTQJKAkMCQgJOAksCRgJJACcA6QBFAQgAYwEpAIoBUACSAVgAmQFfALQBewCxAXgAswF6AMIBiQARANMAEwDVAAoAzAAMAM4ADQDPAA4A0AALAM0ABADGAAYAyAAHAMkACADKAAUAxwA1APcANwD5ADsA/QAtAO8ALwDxADAA8gAxAPMALgDwAE8BEwBNAREAcQE3AHMBOQBqATAAbAEyAG0BMwBuATQAawExAHUBOwB3AT0AeAE+AHkBPwB2ATwAoAFnAKIBaQCkAWsApgFtAKcBbgCoAW8ApQFsALsBggC6AYEAvAGDAL0BhAHsAe0B6AHqAesB6QI8Aj0B1QICAgACPgI2AiMCJwIpAhUCEgIqAh0CHAUjAKIFrQA3Ba0ANwWtADcFrQA3Ba0ANwWtADcFrQA3Ba0ANwWtADcFrQA3Ba0ANwWtADcFrQA3Ba0ANwWtADcFrQA3Ba0ANwWtADcFrQA3Ba0ANwWtADcFrQA3Ba0ANwWtADcFrQA3CCAAAAggAAAFWQCiBZwAYAWcAGAFnABgBZwAYAWcAGAFnABgBaAAogXaAEYFoACiBdoARgWgAKIE9QCiBPUAogT1AKIE9QCiBPUAogT1AKIE9QCiBPUAogT1AKIE9QCiBPUAPwT1AKIE9QCiBPUAogT1AKIE9QCiBPUAogT1AKIE9QCiBPUAogTLAKIF7gBgBe4AYAXuAGAF7gBgBe4AYAYEAKIGcQAiBgQAogYEAKICZwCiAmcAogJn/9YCZ//JAmf+4AJn/8ICZwCiAmcAogJn/6ACZwCiAmf/1AJn/9UCZwA3Amf/4QMkADcDJAA3BZoAogWaAKIEgwCiBIMAogSDAKIEgwCiBIMAogScAAkHKwCiBhwAogYcAKIGHACiBhwAogYcAKIGHACiBhwAogXxAGAF8QBgBfEAYAXxAGAF8QBgBfEAYAXxAGAF8QBgBfEAYAXxAGAF8QBgBfEAYAXxAGAF8QBgBoYAYAaGAGAGhgBgBoYAYAaGAGAGhgBgBfEAYAXxAGAF8QBgBfEAYAYQACsGEAArBfEAYAjOAGAFGQCiBVEAogYKAGAFcACiBXAAogVwAKIFcACiBXAAUwVwAKIFcACiBWgAZAVoAGQFaABkBWgAZAVoAGQFaABkBWgAZAXuAGAE0AAmBNAAJgTQACYE0AAmBNAAJgTQACYFqACiBagAogWoAKIFqACiBagAkQWoAKIFqACiBagAogWoAKIGuQCiBrkAoga5AKIGuQCiBrkAogaxAKIFqACiBagAogWoAKIFqACiBagAogWoAKIFfwAmB5EAJgeRACYHkQAmB5EAJgeRACYFmgA3BTcAEwU3ABMFNwATBTcAEwU3ABMFNwATBTcAEwU3ABMFOwBsBTsAbAU7AGwFOwBsBTsAbARtAFMEbQBTBG0AUwRtAFMEbQBTBG0AUwRtAFMEbQBTBG0AUwRtAFMEbQBTBG0AUwRtAFMEbQBTBG3/9gRtAFMEbQBTBG0AUwRtAFMEbQBTBG0AUwRtAFMEbQBTBG0AUwRtAFMHIwBTByMAUwTBAH0ERgBMBEYATARGAEwERgBMBEYATARGAEwEwQBMBH8ATAS6AEwE/gBMBMEATASBAFMEgQBTBIEAUwRxAFMEgQBTBIEAUwSBAFMEgQBTBIEAUwSBAFMEgf/2BIEAUwSBAFMEgQBTBIEAUwSBAFMEgQBTBIEAUwSBAFMEgQBTBIEAVgMZADIE6QAqBOkAKgTpACoE6QAqBOkAKgShAH0E1QAoBKH/nwShAH0CFgB9AhEAfQIRAH0CEf+uAhH/oQIR/rgDAAANAhEAfQIWAH0CEf/SAhEAfQK0//kCvf/+AhEADAKo//sCQf/gAj//4AI//8MEhwB9BIcAfQSyAH0CkgB9ApIAfQKSAH0CkgB9A1AAfQLmAEYG/wB9BKIAfQSiAH0EogB9BKIAfQSiAH0EsAB9BKIAfQR/AEwEfwBMBH8ATAR/AEwEfwBMBH8ATAR/AEwEfwBMBH8ATAR//+wEfwBMBH8ATAR/AEwEfwBMBQIATAUCAEwFAgBMBQIATAUCAEwFAgBMBH8ATAR/AEwEfwBMBH8ATATHADMExwAzBH8ATAdMAEwEvgB9BLgAfQS+AEwDOAB9AzgAfQM4AFQDOAB9Azj/bAM4AH0DOABgBB4ATQQeAE0EHgBNBB4ATQQeAE0EHgBNBB4ATQTmAH0DQQA8A0EAPAOEADwDQQA8A0EAPANBADwEpgBwBKYAcASmAHAEpgBwBKYAcASm//sEpgBwBKYAcASmAHAEpgBwBYwAcAWMAHAFjABwBYwAcAWMAHAFjABwBKYAcASmAHAEpgBwBKYAcASmAHAEpgBwBBEAGwYpAB4GKQAeBikAHgYpAB4GKQAeBJgAIwRmABsEZgAbBGYAGwRmABsEZgAbBGYAGwRmABsEZgAbA+EAPwPhAD8D4QA/A+EAPwPhAD8FLwAyBasAMgMtAFMDJgBMBUUALgUJAFgDWABWBPYAVgUyAFgFIABWBUYAWAUpAFgE5ABWBT4AWAUmAFwFeACPBXgAuQV4AKYFeACFBXgAcgV4AJYFeACABXgAewV4AHUFeACMAyQANAIBACoDEAAqA1IAPgM0ACoDSwA0AzgANALpACoDWwA+AzcANwMkADQCAQAqAxAAKgNSAD4DNAAqA0sANAM4ADQC6QAqA1sAPgM3ADcDJAA0AgEAKgMQACoDUgA+AzQAKgNLADQDOAA0AukAKgNbAD4DNwA3Adr+sgcLACoHNwAqCHgAKgZ+ACoHfwA+B0AAKghBAD4IOgA0B5EAKgIbAGMCKgBeAhsAYwItAGMG/QBjAjQAZQI1AGUELQAWBD8AQwIGAGMCwgBcA7IAOQTQADgC+gAXAvoAFwJYAIwCWACRApoAdwKhABgDAQAYAwAAHQKiAJQCogBEAxIAeAMSAHgEoQB4CPkBaARSAHQGKAAAAioAXgPlAF4DvwBNA7wAZAIOAE0CCwBkA+cAZQP+AHoCagBlAmkAfAO/AHACCgBwAcMAAAHDAAAB3AAABGIAZwWcAGAEhABLBVEAYwQwAEwFFwABBA7/yQTh/9oFEAApBR8AHgZ9ABEFBgApBRAAKQi9ABcFJv/1BXgA7wV4AEUFeAB2BXgA3QV4AFMFeACRBXgAZgV4AHEFeABSBXgAZwV4AHEFeAAeAsQAvgTEAFgEmQB0BJkAdASBAHQEmQB0BJkAdASZAHQEgwB0BIMAdASZAHQEmQB0BJkAdATBAHQEyABxBJkAdAPIAEUEqwBSBjEAXQOqABgF/wBcBgYAIQV0ACUEeQA5BVoAHgS7AFMEuABwB1YAXQqRAF0EWAB2BncATAYfAF4ErgBMBAEAOAZ9AFEEUQBRBn0AUQdyABUDtACFAhkAogIZAKIEdAArBBEAPgQRAD4JQgCiB+YAPAH9AGQB/QBkAAAACgAAAAoAAAAL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SGJhcgtIY2lyY3VtZmxleAd1bmkxRTI0BklicmV2ZQd1bmkwMjA4B3VuaTFFQ0EHdW5pMUVDOAd1bmkwMjBBB0ltYWNyb24HSW9nb25lawZJdGlsZGULSmNpcmN1bWZsZXgHdW5pMDEzNgZMYWN1dGUGTGNhcm9uB3VuaTAxM0IETGRvdAZOYWN1dGUGTmNhcm9uB3VuaTAxNDUHdW5pMUU0NANFbmcGT2JyZXZlB3VuaTFFRDAHdW5pMUVEOAd1bmkxRUQyB3VuaTFFRDQHdW5pMUVENgd1bmkwMjBDB3VuaTFFQ0MHdW5pMUVDRQVPaG9ybgd1bmkxRURBB3VuaTFFRTIHdW5pMUVEQwd1bmkxRURFB3VuaTFFRTANT2h1bmdhcnVtbGF1dAd1bmkwMjBFB09tYWNyb24HdW5pMDFFQQtPc2xhc2hhY3V0ZQZSYWN1dGUGUmNhcm9uB3VuaTAxNTYHdW5pMDIxMAd1bmkxRTVBB3VuaTAyMTIGU2FjdXRlC1NjaXJjdW1mbGV4B3VuaTAyMTgHdW5pMUU2Mgd1bmkwMThGBFRiYXIGVGNhcm9uB3VuaTAxNjIHdW5pMDIxQQd1bmkxRTZDBlVicmV2ZQd1bmkwMjE0B3VuaTFFRTQHdW5pMUVFNgVVaG9ybgd1bmkxRUU4B3VuaTFFRjAHdW5pMUVFQQd1bmkxRUVDB3VuaTFFRUUNVWh1bmdhcnVtbGF1dAd1bmkwMjE2B1VtYWNyb24HVW9nb25lawVVcmluZwZVdGlsZGUGV2FjdXRlC1djaXJjdW1mbGV4CVdkaWVyZXNpcwZXZ3JhdmULWWNpcmN1bWZsZXgHdW5pMUVGNAZZZ3JhdmUHdW5pMUVGNgd1bmkxRUY4BlphY3V0ZQpaZG90YWNjZW50B3VuaTFFOTIGYWJyZXZlB3VuaTFFQUYHdW5pMUVCNwd1bmkxRUIxB3VuaTFFQjMHdW5pMUVCNQd1bmkxRUE1B3VuaTFFQUQHdW5pMUVBNwd1bmkxRUE5B3VuaTFFQUIHdW5pMDIwMQd1bmkxRUExB3VuaTFFQTMHdW5pMDIwMwdhbWFjcm9uB2FvZ29uZWsKYXJpbmdhY3V0ZQdhZWFjdXRlC2NjaXJjdW1mbGV4CmNkb3RhY2NlbnQGZGNhcm9uB3VuaTFFMEQGZWJyZXZlBmVjYXJvbgd1bmkxRUJGB3VuaTFFQzcHdW5pMUVDMQd1bmkxRUMzB3VuaTFFQzUHdW5pMDIwNQplZG90YWNjZW50B3VuaTFFQjkHdW5pMUVCQgd1bmkwMjA3B2VtYWNyb24HZW9nb25lawd1bmkxRUJEB3VuaTAyNTkLZ2NpcmN1bWZsZXgHdW5pMDEyMwpnZG90YWNjZW50BGhiYXILaGNpcmN1bWZsZXgHdW5pMUUyNQZpYnJldmUHdW5pMDIwOQlpLmxvY2xUUksHdW5pMUVDQgd1bmkxRUM5B3VuaTAyMEIHaW1hY3Jvbgdpb2dvbmVrBml0aWxkZQd1bmkwMjM3C2pjaXJjdW1mbGV4B3VuaTAxMzcMa2dyZWVubGFuZGljBmxhY3V0ZQZsY2Fyb24HdW5pMDEzQwRsZG90Bm5hY3V0ZQZuY2Fyb24HdW5pMDE0Ngd1bmkxRTQ1A2VuZwZvYnJldmUHdW5pMUVEMQd1bmkxRUQ5B3VuaTFFRDMHdW5pMUVENQd1bmkxRUQ3B3VuaTAyMEQHdW5pMUVDRAd1bmkxRUNGBW9ob3JuB3VuaTFFREIHdW5pMUVFMwd1bmkxRUREB3VuaTFFREYHdW5pMUVFMQ1vaHVuZ2FydW1sYXV0B3VuaTAyMEYHb21hY3Jvbgd1bmkwMUVCC29zbGFzaGFjdXRlBnJhY3V0ZQZyY2Fyb24HdW5pMDE1Nwd1bmkwMjExB3VuaTFFNUIHdW5pMDIxMwZzYWN1dGULc2NpcmN1bWZsZXgHdW5pMDIxOQd1bmkxRTYzBHRiYXIGdGNhcm9uB3VuaTAxNjMHdW5pMDIxQgd1bmkxRTZEBnVicmV2ZQd1bmkwMUQ0B3VuaTAyMTUHdW5pMUVFNQd1bmkxRUU3BXVob3JuB3VuaTFFRTkHdW5pMUVGMQd1bmkxRUVCB3VuaTFFRUQHdW5pMUVFRg11aHVuZ2FydW1sYXV0B3VuaTAyMTcHdW1hY3Jvbgd1b2dvbmVrBXVyaW5nBnV0aWxkZQZ3YWN1dGULd2NpcmN1bWZsZXgJd2RpZXJlc2lzBndncmF2ZQt5Y2lyY3VtZmxleAd1bmkxRUY1BnlncmF2ZQd1bmkxRUY3B3VuaTFFRjkGemFjdXRlCnpkb3RhY2NlbnQHdW5pMUU5Mwd6ZXJvLnRmBm9uZS50ZgZ0d28udGYIdGhyZWUudGYHZm91ci50ZgdmaXZlLnRmBnNpeC50ZghzZXZlbi50ZghlaWdodC50ZgduaW5lLnRmCXplcm8uZG5vbQhvbmUuZG5vbQh0d28uZG5vbQp0aHJlZS5kbm9tCWZvdXIuZG5vbQlmaXZlLmRub20Ic2l4LmRub20Kc2V2ZW4uZG5vbQplaWdodC5kbm9tCW5pbmUuZG5vbQl6ZXJvLm51bXIIb25lLm51bXIIdHdvLm51bXIKdGhyZWUubnVtcglmb3VyLm51bXIJZml2ZS5udW1yCHNpeC5udW1yCnNldmVuLm51bXIKZWlnaHQubnVtcgluaW5lLm51bXIHdW5pMjA3MAd1bmkwMEI5B3VuaTAwQjIHdW5pMDBCMwd1bmkyMDc0B3VuaTIwNzUHdW5pMjA3Ngd1bmkyMDc3B3VuaTIwNzgHdW5pMjA3OQd1bmkyMTUzB3VuaTIxNTQJb25lZWlnaHRoDHRocmVlZWlnaHRocwtmaXZlZWlnaHRocwxzZXZlbmVpZ2h0aHMJcGVyaW9kLnRmCGNvbW1hLnRmB3VuaTAwQUQKZW1kYXNoLmFsdAd1bmkwMEEwAkNSDWNvbG9ubW9uZXRhcnkEZG9uZwRFdXJvBGxpcmEHdW5pMjBCQQd1bmkyMEE2B3VuaTIwQjkHdW5pMjBBOQdjZW50LnRmEGNvbG9ubW9uZXRhcnkudGYJZG9sbGFyLnRmB2RvbmcudGYHRXVyby50ZglmbG9yaW4udGYIZnJhbmMudGYHbGlyYS50Zgp1bmkyMEJBLnRmCnVuaTIwQjkudGYLc3RlcmxpbmcudGYGeWVuLnRmB3VuaTIyMTkHdW5pMjIxNQhlbXB0eXNldAd1bmkyMTI2B3VuaTIyMDYHdW5pMDBCNQd1bmkyMTE3B3VuaTIxMTMHdW5pMjExNgd1bmkyMTIwB3VuaTAyQkMHdW5pMDJCQgd1bmkwMzA4B3VuaTAzMDcJZ3JhdmVjb21iCWFjdXRlY29tYgd1bmkwMzBCC3VuaTAzMEMuYWx0B3VuaTAzMDIHdW5pMDMwQwd1bmkwMzA2B3VuaTAzMEEJdGlsZGVjb21iB3VuaTAzMDQNaG9va2Fib3ZlY29tYgd1bmkwMzBGB3VuaTAzMTEHdW5pMDMxMgd1bmkwMzFCDGRvdGJlbG93Y29tYgd1bmkwMzI2B3VuaTAzMjcHdW5pMDMyOBBncmF2ZWNvbWIubmFycm93EGFjdXRlY29tYi5uYXJyb3cLdW5pMDMwNjAzMDELdW5pMDMwNjAzMDALdW5pMDMwNjAzMDkLdW5pMDMwNjAzMDMLdW5pMDMwMjAzMDELdW5pMDMwMjAzMDALdW5pMDMwMjAzMDkLdW5pMDMwMjAzMDMHdW5pMDAwMAAAAAABAAAACgAoAFAAAkRGTFQADmxhdG4ADgAEAAAAAP//AAMAAAABAAIAA2tlcm4AIm1hcmsAGm1rbWsAFAAAAAEAAwAAAAIAAQACAAAAAQAAAAQRPBCwACQACgAGABAAAQAKAAAAAREiESIAAREWAAwAAREQAAQAAAABAAgAARBuDuoABA9+AAwBfQ7YAAAO0g7MDtgAAA7SDswO2AAADtIOzA7YAAAO0g7MDtgAAA7SDswO2AAADtIOzA7YAAAO0g7MDtgAAA7SDswO2AAADtIOzA7YAAAO0g7MDtgAAA7SDswO2AAADtIOzA7YAAAO0g7MDtgAAA7SDswO2AAADtIOzA7YAAAO0g7MDtgAAA7SDswO2AAADtIOzA7YAAAO0g7MDtgAAA7SDswO2AAADtIOzA7YAAAO0g7MDtgAAA7SDswO2AAADtIOzA7YAAAO0g7MDsYAAAAAAAAOxgAAAAAAAA7AAAAOugAADsAAAA66AAAOwAAADroAAA7AAAAOugAADsAAAA66AAAOwAAADroAAA60AAAOrgAADqgAAA6iAAAOtAAADq4AAA6oAAAOogAADrQAAA6uAAAOnAAADpYOkA6cAAAOlg6QDpwAAA6WDpAOnAAADpYOkA6cAAAOlg6QDpwAAA6WDpAOnAAADpYOkA6cAAAOlg6QDpwAAA6WDpAOnAAADpYOkA6cAAAOlg6QDpwAAA6WDpAOnAAADpYOkA6cAAAOlg6QDpwAAA6WDpAOnAAADpYOkA6cAAAOlg6QDpwAAA6WDpAOnAAADpYOkA6cAAAOlg6QDooAAA6EAAAOigAADoQAAA6KAAAOhAAADooAAA6EAAAOigAADoQAAA5+AAAOeAAADnIAAA5sAAAOfgAADngAAA5+AAAOeAAADmYAAA5gDloOZgAADmAOWg5mAAAOYA5aDmYAAA5gDloOZgAADmAOWg5mAAAOYA5aDmYAAA5gDloOZgAADmAOWg5mAAAOYA5aDmYAAA5gDloOZgAADmAOWg5mAAAOYA5aDmYAAA5gDloOZgAADmAOWg5UAAAAAAAADlQAAAAAAAAOTgAADkgAAA5OAAAOSAAADkIAAA48AAAOQgAADjwAAA5CAAAOPAAADkIAAA48AAAOQgAADjwAAA42AAAOMAAADioAAA4kAAAOKgAADiQAAA4qAAAOJAAADioAAA4kAAAOKgAADiQAAA4qAAAOJAAADioAAA4kAAAOHg4YDhIODA4eDhgOEg4MDh4OGA4SDgwOHg4YDhIODA4eDhgOEg4MDh4OGA4SDgwOHg4YDhIODA4eDhgOEg4MDh4OGA4SDgwOHg4YDhIODA4eDhgOEg4MDh4OGA4SDgwOHg4YDhIODA4eDhgOEg4MDh4OGA4SDgwOHg4YDhIODA4eDhgOEg4MDh4OGA4SDgwOHg4YDhIODA4eDhgOEg4MDh4OGA4SDgwOHg4YDhIODA4eDhgOEg4MDh4OGA4SDgwOBgAAAAAAAA4GAAAAAAAADh4OGA4SDgwOAAAAAAAAAA36AAAN9AAADfoAAA30AAAN+gAADfQAAA36AAAN9AAADfoAAA30AAAN+gAADfQAAA36AAAN9AAADe4AAA3oAAAN7gAADegAAA3uAAAN6AAADe4AAA3oAAAN7gAADegAAA3uAAAN6AAADe4AAA3oAAAN4gAADdwAAA3iAAAN3AAADeIAAA3cAAAN4gAADdwAAA3iAAAN3AAADeIAAA3cAAAN1g3QDcoNxA3WDdANyg3EDdYN0A3KDcQN1g3QDcoNxA3WDdANyg3EDdYN0A3KDcQN1g3QDcoNxA3WDdANyg3EDdYN0A3KDcQN1g3QDcoNxA3WDdANyg3EDdYN0A3KDcQN1g3QDcoNxA3WDdANyg3EDdYN0A3KDcQN1g3QDcoNxA3WDdANyg3EDdYN0A3KDcQN1g3QDcoNxA3WDdANyg3EDdYN0A3KDcQNvgAAAAAAAA2+AAAAAAAADb4AAAAAAAANvgAAAAAAAA2+AAAAAAAADbgAAA2yAAANuAAADbIAAA24AAANsgAADbgAAA2yAAANuAAADbIAAA24AAANsgAADbgAAA2yAAANuAAADbIAAA2sAAANpgAADawAAA2mAAANrAAADaYAAA2sAAANpgAADawAAA2mAAANoAAADZoNlA2gAAANmg2UDaAAAA2aDZQNoAAADZoNlA2gAAANmg2UDaAAAA2aDZQNoAAADZoNlA2gAAANmg2UDaAAAA2aDZQNoAAADZoNlA2gAAANmg2UDaAAAA2aDZQNoAAADZoNlA2gAAANmg2UDaAAAA2aDZQNoAAADZoNlA2gAAANmg2UDaAAAA2aDZQNoAAADZoNlA2gAAANmg2UDaAAAA2aDZQNoAAADZoNlA2gAAANmg2UDaAAAA2aDZQNoAAADZoNlA2OAAANmg2UDY4AAA2aDZQAAAAADYgAAA2CAAANfAAADYIAAA18AAANggAADXwAAA2CAAANfAAADYIAAA18AAANggAADXwAAAAAAAANdgAAAAAAAA12AAAAAAAADXYAAAAAAAANdgAADXAAAA1qDWQNcAAADWoNZA1wAAANag1kDXAAAA1qDWQNcAAADWoNZA1wAAANag1kDXAAAA1qDWQNcAAADWoNZA1wAAANag1kDXAAAA1qDWQNcAAADWoNZA1wAAANag1kDXAAAA1qDWQNcAAADWoNZA1wAAANag1kDXAAAA1qDWQNcAAADWoNZA1wAAANag1kDXAAAA1qDWQNcAAADWoNZA1eAAANWA1SDUwAAAAAAAANTAAAAAAAAA1MAAAAAAAADUwAAAAAAAANTAAAAAAAAA1GAAANQAAADToAAA00AAANRgAADUAAAA1GAAANQAAAAAAAAA0uAAANKAAAAAANIg0oAAAAAA0iDSgAAAAADSINKAAAAAANIg0oAAAAAA0iDRwAAAAADRYNKAAAAAANIgAAAAANLgAADSgAAAAADSINKAAAAAANIg0QAAAAAA0KDQQAAAAADP4NKAAAAAANIgz4AAAAAAzyDOwAAAAAAAAM7AAAAAAAAAAAAAAM5gAAAAAAAAzmAAAM4AAADNoAAAzgAAAM2gAADOAAAAzaAAAM4AAADNoAAAzgAAAM2gAADNQAAAzOAAAMyAAADMIAAAzIAAAMwgAADMgAAAzCAAAMyAAADMIAAAzIAAAMwgAADMgAAAzCAAAMvAy2DLAMqgy8DLYMsAyqDLwMtgywDKoMvAy2DLAMqgy8DLYMsAyqDLwMtgywDKoMvAy2DLAMqgy8DLYMsAyqDLwMtgywDKoMvAy2DLAMqgy8DLYMsAyqDLwMtgywDKoMvAy2DLAMqgy8DLYMsAyqDLwMtgywDKQMvAy2DLAMpAy8DLYMsAykDLwMtgywDKQMvAy2DLAMpAy8DLYMsAykDLwMtgywDKoMvAy2DLAMqgy8DLYMsAyqDLwMtgywDKoMngyYDJIMjAyeDJgMkgyMDLwMtgywDKoMhgAAAAAAAAyAAAAMegAADIAAAAx6AAAMgAAADHoAAAyAAAAMegAADIAAAAx6AAAMgAAADHoAAAyAAAAMegAADHQAAAxuAAAMdAAADG4AAAx0AAAMbgAADHQAAAxuAAAMdAAADG4AAAx0AAAMbgAADHQAAAxuAAAAAAAADGgAAAAAAAAMaAAAAAAAAAxoAAAAAAAADGgAAAAAAAAMaAAAAAAAAAxoAAAMYgxcDFYMUAxiDFwMVgxQDGIMXAxWDFAMYgxcDFYMUAxiDFwMVgxQDGIMXAxWDFAMYgxcDFYMUAxiDFwMVgxQDGIMXAxWDFAMYgxcDFYMUAxiDFwMVgxQDGIMXAxWDFAMYgxcDFYMUAxiDFwMVgxQDGIMXAxWDFAMYgxcDFYMUAxiDFwMVgxQDGIMXAxWDFAMYgxcDFYMUAxiDFwMVgxQDGIMXAxWDFAMYgxcDFYMUAxKAAAAAAAADEoAAAAAAAAMSgAAAAAAAAxKAAAAAAAADEoAAAAAAAAMRAAAAAAAAAxEAAAAAAAADEQAAAAAAAAMRAAAAAAAAAxEAAAAAAAADEQAAAAAAAAMRAAAAAAAAAxEAAAAAAAADD4AAAw4AAAMPgAADDgAAAw+AAAMOAAADD4AAAw4AAAMPgAADDgAAAwyAAAMLAAADsAAAA66AAAMJgAADCAAAAwaAAAAAAAADBQAAAwOAAAMCAAADAIAAAv8AAAL9gAADioAAA4kAAAL8AAAC+oAAAABAa0AAAABAbADiAABAYsAAQABAYsDiAABAtYAAAABAtEFpgABAuIA3gABAuEE6AABBE4FpgABAz8AAAABA04FpgABAloA3gABAlkE6AABAfQAAAABAfQECgABAjQECgABAx0ECgABA/8AFQABAlQAAAABA/oC9gABAlAECgABAeYADAABAhcAAAABAiAECgABAQ4AAAABAcEECgABA7QECgABAwsAFgABAm4AAAABA8kC9gABAl4ECgABAtAAFAABAtoAFgABAj0AAAABA5gC9gABAkEECgABAlcAAAABAl0ECgABAVEAAQABAVQF+AABAQgAAQABAQwF+AABAkcAAAABAS8ECgABAbgAFQABAU8ECgABAccAFQABAV4ECgABAcMAFQABAVoECgABAeYAFQABAXwECgABAXYAFQABAQ0ECgABAQsAAAABApMAAAABAUAFpgABAl4AAAABAQwFpgABAmEECgABAUwD3AABAjYAAAABAjMECgABAzQALgABAk4AAAABAkoECgABAiAAAAABAj4AAAABAj4ECgABArcAAAABA2cECgABA9MAFwABAkYAAAABAksECgABArYAAAABAtIFpgABApwAAAABAr0FpwABA8gFpgABA3wAFgABAuMAAQABBO4EawABAuUFpgABAnsAAQABAnsFpgABArIAAAABArYFpgABAroAAAABAqgFpgABBFsFpgABAwEFpgABA6MAEwABAv0AAAABBMEEawABAvoFpgABAv8AAAABAw4FpgABAoAAAAABAUgFpgABAmcAAAABATAFpgABAr4AAAABAtoFpgABAgoFpgABAaIADAABAUMAAAABATUFpgABA0gAAAABA0EFqgABAwkAAAABAwMFqgABAu0AAAABAvMFpgABBIkAFAABAqUAAAABApQFpgABAv4AAAABAwQFpgABAsQAAAABAsoFpgABAvEAAAABAu0FpgABBTgFpgABBVUADgABAtcAAAABAtcFpgACABgAAQAbAAAAHQA7ABsAPQBdADoAXwCBAFsAhQCSAH4AlACuAIwAsAC0AKcAtgDlAKwA5wD+ANwBAAEXAPQBGQEcAQwBHgEjARABJQEpARYBKwFHARsBSwFYATgBWgF1AUYBdwF7AWIBfQGJAWcB9wH4AXQCAQIBAXYCBAIEAXcCBgIHAXgCNwI3AXoCPgI/AXsAHAAAAOoAAADkAAAA3gAAANgAAADSAAAA6gAAAOoAAADMAAAAxgAAAMAAAAC6AAAAtAAAAK4AAACoAAAAogABAJwAAgCWAAIAkAACAIoAAwCEAAAAfgAAAH4AAAB+AAAAfgAAAOoAAADqAAAAeAAAAHIAAQF7BAoAAQF3BAoAAQF8BAsAAQGNABUAAQENAAAAAQCNAAAAAQCVAAAAAQAAA48AAQDuBAoAAQF1BAoAAQJfBAoAAQJHBBIAAQGCBAoAAQGXBAoAAQD8A/sAAQF+BAsAAQEjBAoAAQB3BAoAAQGfBAoAAQCWBAoAAQF9BAoAAgADAkICRgAAAkgCVgAFAmYCbQAU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bf2TAP//AKIAAAViByYCJgBCAAAABwJIAYUBoP//AKL+cgViBaYCJgBCAAAABwJTAnUAAAABAKIAAAHFBaYAAwAAMxEhEaIBIwWm+loA//8AogAAAsIHQgImAEYAAAAHAkUAvQGc////1gAAApMHHQImAEYAAAAHAkr/twGb////yQAAApwHIgImAEYAAAAHAkj/twGc///+4AAAAksHQgImAEYAAAAHAk/+1gGc////wgAAAqgHQgImAEYAAAAHAkL/uAGc//8AogAAAcUHQgImAEYAAAAHAkMAnwGc//8Aov5yAc8FpgImAEYAAAAHAlMArgAA////oAAAAcUHQgImAEYAAAAHAkT/lgGc//8AogAAAhoHtgImAEYAAAAHAk7+7QGU////1AAAApIHNwImAEYAAAAHAlD/wAGc////1QAAApQHDAImAEYAAAAHAk3/sgGc//8AN/6mAeMFpgImAEYAAAAGAlYV9////+EAAAKTBz8CJgBGAAAABwJM/54BnAABADf/7AKiBaYAEAAABSImJxMWFjMyNjURIREUBgYBMEWSIgYdSy5lRgEkPKAUFwkBCwYUa3gDxvvsirxg//8AN//sA3AHIgImAFQAAAAHAkgAjAGcAAEAogAABWQFpgALAAAzESERASEBASEBAxGiAS8CDwFP/kkB7P7C/oHWBab9agKW/dD8igKh/v7+Yf//AKL+SAVkBaYCJgBWAAAABwJUAjEAAAABAKIAAAQ8BaYABQAAMxEhESEVogEmAnQFpvtR9///AKIAAAQ8B0ICJgBYAAAABwJFALgBnP//AKIAAAQ8BaYCJgBYAAAABwJHAqcAAP//AKL+SAQ8BaYCJgBYAAAABwJUAdoAAP//AKIAAAQ8BaYCJgBYAAAABwHUAn0ASgACAAkAAARUBaYAAwAJAAATNSUVARMhESEVCQMN/aQBASUCdAI48s7n/O8FpvtR9wABAKIAAAaKBaYADAAAMxEhAQEhESERASMBEaIBqAFLAU4Bp/7Z/qTl/qYFpvv5BAf6WgRS+64ETfuzAAABAKIAAAV7BaYACQAAMxEhAREhESEBEaIBIQKUAST++f1RBab8QwO9+loD4/wdAP//AKIAAAV7B0ICJgBfAAAABwJFApYBnP//AKIAAAV7ByQCJgBfAAAABwJJAZABnP//AKL+SAV7BaYCJgBf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';
    const FONTE_PUBLIC_SANS_ITALIC = 'AAEAAAAQAQAABAAAR0RFRjApJ6EAAAI8AAAA7kdQT1PMTtLeAAAx0AAAMXRHU1VCs1+0hgAABrAAAASwT1MvMpEJZH8AAAHcAAAAYFNUQVTnccwZAAABmAAAAERjbWFwsBMzOQAAEEAAAAZUZ2FzcAAAABAAAAEUAAAACGdseWZTGkp0AABjRAAAgHxoZWFkIFRGtAAAAWAAAAA2aGhlYQ9vCn8AAAE8AAAAJGhtdHj7g9KyAAAWlAAACbxsb2NhRpBnYAAAC2AAAATgbWF4cAJ/ANEAAAEcAAAAIG5hbWVXxnJuAAADLAAAA4Jwb3N0CdkYUAAAIFAAABF9cHJlcGgGjIUAAAEMAAAAB7gB/4WwBI0AAAEAAf//AA8AAQAAAm8AagAHAGUABQABAAAAAAAAAAAAAAAAAAMAAQABAAAHbP4+AAAJc/59/HcI9gPoAI0AAAAAAAAAAAAAAAACbwABAAAAAgBC0YXHsF8PPPUAAwfQAAAAANs01VYAAAAA3pIl/P59/iAI9gg1AAIABgACAAAAAAAAAAEAAQAIAAIAAAAUAAIAAAAkAAJ3Z2h0AQAAAGl0YWwBBAABABAABAABAAEAAAEEAAEAAAADAAAAAgEWAZAAAAK8AAAABARRAZAABQAABRQEsP/rAJYFFASwAGICvACMAmwAAAAAAAAAAAAAAACgAAD/QAAgWwAAAAAAAAAATk9ORQCBAAD7Agds/j4AAAkLAe8gAAGTAAAAAAQKBaYAAAAgAAMAAQACAC4AAAAOAAAA4AAIAAIAGAAQAAEAAgGKAYsAAQAEAAECpQABAAQAAQJyAAIAHQABABsAAQAdADsAAQA9AF0AAQBfAIEAAQCFAJIAAQCUAK4AAQCwALQAAQC2AN0AAQDfAOUAAQDnAP4AAQEAARcAAQEZARwAAQEeASMAAQElASkAAQErAUcAAQFLAVgAAQFaAXUAAQF3AXsAAQF9AYkAAQGKAYsAAgH3AfgAAQIBAgEAAQIEAgQAAQIGAgcAAQI3AjcAAQI+Aj8AAQJCAkYAAwJIAlYAAwJmAm0AAwABAAEAAAAIAAEAAQJSAAAAAAALAIoAAwABBAkAAACqAk4AAwABBAkAAQAWAjgAAwABBAkAAgAMAiwAAwABBAkAAwA4AfQAAwABBAkABAAkAdAAAwABBAkABQAaAbYAAwABBAkABgAiAZQAAwABBAkADgF6ABoAAwABBAkBAAAMAA4AAwABBAkBBAAMAiwAAwABBAkBFgAOAAAAUgBlAGcAdQBsAGEAcgBXAGUAaQBnAGgAdABTAEkATAAgAE8AcABlAG4AIABGAG8AbgB0ACAATABpAGMAZQBuAHMAZQAsACAAVgBlAHIAcwBpAG8AbgAgADEALgAxADoAIABoAHQAdABwAHMAOgAvAC8AcwBjAHIAaQBwAHQAcwAuAHMAaQBsAC4AbwByAGcALwBjAG0AcwAvAHMAYwByAGkAcAB0AHMALwBwAGEAZwBlAC4AcABoAHAAPwBzAGkAdABlAF8AaQBkAD0AbgByAHMAaQAmAGkAZAA9AE8ARgBMAF8AdwBlAGIAOwAgAFUAUwBXAEQAUwAgAE0AbwBkAGkAZgBpAGUAZAAgAFYAZQByAHMAaQBvAG4AOgAgAGgAdAB0AHAAcwA6AC8ALwBnAGkAdABoAHUAYgAuAGMAbwBtAC8AdQBzAHcAZABzAC8AcAB1AGIAbABpAGMALQBzAGEAbgBzAC8AYgBsAG8AYgAvAG0AYQBzAHQAZQByAC8ATABJAEMARQBOAFMARQAuAG0AZABQAHUAYgBsAGkAYwBTAGEAbgBzAC0ASQB0AGEAbABpAGMAVgBlAHIAcwBpAG8AbgAgADIALgAwADAAMQBQAHUAYgBsAGkAYwAgAFMAYQBuAHMAIABJAHQAYQBsAGkAYwAyAC4AMAAwADEAOwBOAE8ATgBFADsAUAB1AGIAbABpAGMAUwBhAG4AcwAtAEkAdABhAGwAaQBjAEkAdABhAGwAaQBjAFAAdQBiAGwAaQBjACAAUwBhAG4AcwBDAG8AcAB5AHIAaQBnAGgAdAAgADIAMAAxADUAIABUAGgAZQAgAFAAdQBiAGwAaQBjACAAUwBhAG4AcwAgAFAAcgBvAGoAZQBjAHQAIABBAHUAdABoAG8AcgBzACAAKABoAHQAdABwAHMAOgAvAC8AZwBpAHQAaAB1AGIALgBjAG8AbQAvAHUAcwB3AGQAcwAvAHAAdQBiAGwAaQBjAC0AcwBhAG4AcwApAAAAAQAAAAoAxAFiAAJERkxUAKBsYXRuAA4AfAAIQVpFIACWQ0FUIABkQ1JUIACWS0FaIACWTU9MIABMUk9NIACWVEFUIACWVFJLIAA0AAD//wAJAAAAAQAD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AToAowFqAH0BQwAYANoAGwDdAH8BRQAPANEAFADWADIA9AA4APoASgEOAFABFABvATUAewFBAIkBTwCLAVEAngFlAKoBcQCRAVcAmAFeAkECQAJeAlgCXwJjAmACWwJEAkUCSAJMAk0CSgJDAkICTgJLAkYCSQAnAOkARQEIAGMBKQCKAVAAkgFYAJkBXwC0AXsAsQF4ALMBegDCAYkAEQDTABMA1QAKAMwADADOAA0AzwAOANAACwDNAAQAxgAGAMgABwDJAAgAygAFAMcANQD3ADcA+QA7AP0ALQDvAC8A8QAwAPIAMQDzAC4A8ABPARMATQERAHEBNwBzATkAagEwAGwBMgBtATMAbgE0AGsBMQB1ATsAdwE9AHgBPgB5AT8AdgE8AKABZwCiAWkApAFrAKYBbQCnAW4AqAFvAKUBbAC7AYIAugGBALwBgwC9AYQB7AHtAegB6gHrAekCPAI9AdUCAgIAAj4CNgIjAicCKQIVAhICKgIdAhwGZwFZBVIADQVSAA0FUgANBVIADQVSAA0FUgANBVIADQVSAA0FUgANBVIADQVSAA0FUgANBVIADQVSAA0FUgANBVIADQVSAA0FUgANBVIADQVSAA0FUgANBVIADQVSAA0FUgANBVIADQdE/7cHRP+3BQkAdgU0AHUFNAB1BTQAdQU0AHUFNAB1BTQAdQU+AHgFhQBgBT4AeAWFAGAFPgB4BLIAdgSyAHYEsgB2BLIAdgSyAHYEsgB2BLIAdgSyAHYEsgB2BLIAdgSyAHYEsgB2BLIAdgSyAHYEsgB2BLIAdgSyAHYEsgB2BLIAdgSyAHYEfQB3BYYAagWGAGoFhgBqBYYAagWGAGoFxwB2BhQAbgXHAHYFxwB2AjkAdwI5AHcCOQB3AjkAdwI5/+cCOQB3AjkAdwI5AE0COQBzAjkAdwI5AHcCOQB3Ajn/wgI5AHcC9wARAvcAEQUjAHYFIwB2BFoAdgRaAHYEWgB2BFoAdgR+AHYEhABKBqoAdwWzAHYFswB2BbMAdgWzAHYFswB2BbMAdgWzAHYFggBqBYIAagWCAGoFggBqBYIAagWCAGoFggBqBYIAagWCAGoFggBqBYIAagWCAGoFggBqBYIAagX3AGoF9wBqBfcAagX3AGoF9wBqBfcAagWCAGoFggBqBYIAagWCAGoFmf/6BZn/+gWCAGoIVwBqBMsAdwToAHYFnwBqBRkAdgUZAHYFGQB2BRkAdgUZAHYFGQB2BRkAdgUeAGQFHgBkBR4AZAUeAGQFHgBkBR4AZAUeAGQFhwBwBHAAmwRwAJsEcACbBHAAmwRwAJsEcACbBUcApgVHAKYFRwCmBUcApgVHAKYFRwCmBUcApgVHAKYFRwCmBhYApgYTAKYGEwCmBhMApgYTAKYGDwCmBTYApgVHAKYFRwCmBUcApgVHAKYFRwCmBPgAswcWALMHFgCzBxYAswcWALMHFgCzBS4ADQSGAJsEhgCbBIYAmwSGAJsEhgCbBIYAmwSGAJsEhgCbBMwANgTMADYEzAA2BMwANgTMADYEKQA4BEgAOARIADgESAA4BCkAOARIADgESAA4BEgAOARIADgEMQA4BEgAOAQ8ADgERgA4BEgAOAR9AG0ESAA4BEgAOARMADwESAA4BEgAOARIADgEKAA4BEgAOAQpADgESAA4BtsANwbbADcEhQBaBBQARAQNAEQEFABEBBQARAQUAEQEFABEBIIAOgRCAFcEaAA6BIoAOgSCADoEQABNBEAATQRAAE0EOQBNBEAATQRAAE0EQABNBDEATQRAAE0EQABNBEAATQRAAE0EQABNBEAATQRAAE0EQABNBEAATQRAAE0EQABNBEAATQRAAE0C+QB+BKX/6ASl/+gEpf/oBKX/6ASl/+gEZwBaBMsAiQRnAFoEZwBaAesAVQHqAFAB6gBQAeoAUAHqACUB6v+FAeoAUAHqAFAB7AAgAeoAUAHmAFAB6gAtAeoARQHr/5cB6gBPAgL/XQH6/10B7f9dBDkAUAQ5AFAESwBQAlEAbgJYAG4CUQBuAlEAFgLKAG4C3QBGBnsAUARaAFAEWgBQBFoAUARaAFAEWgBQBE8AUARaAFAEPQBFBD0ARQQ9AEUEPQBFBB0ARQQ9AEUEPQBFBD0ARQQ9AEUEPQBFBD0ARQQ9AEUEPQBFBD0ARQSVAEUElQBFBJUARQSkAFQElQBFBJUARQQXAEUEPQBFBD0ARQQ9AEUEI//pBCP/6QQ9AEUHEABFBIAALARwACMEggBFAvMAUALqAFAC8wBQAvP//wLzAAgC8wAVAvMAUAPRAEAD3QBAA8oAQAPRAEAD0QBAA9EAQAPRAEAEUgBRAvMAkgLlAGwDCwCSAvMAcwLzAJIC8wCSBGQAYwRkAGMEZABjBGQAYwRkAGMEZABjBGQAYwRkAGMEZABjBGQAYwTjAGME2QBjBOMAYwTtAG0E4wBjBOMAYwQnAGMEZABjBGQAYwRkAGMEZABjBGQAYwPDAHgFywB9BcsAfQXLAH0FywB9BcsAfQQ3//YD/f/nA/3/5wP9/+cD/f/nA/3/5wP9/+cD/f/nA/3/5wOfAAYDyAAGA7YABgOfAAYDnwAGBOQAfgVKAH4DAgCaAvoAnwUtAHIEmgBKAxkAtwR8ADIEvQBDBLEASATQAEkErgBbBH8AyQTdAD0EsgBLBXgAuQV4AJ0FeACyBXgAowV4AJkFeACPBXgAuQV4AP4FeACQBXgAvQLVAAwBzgBLArv/9wMAAA8C7QAFAvgADQLiAA0CpwBXAxQACgLkAA0C1gBcAc0AmAK7AEYDAABeAu0AUwL4AFsC4gBdAqYApAMUAFkC5ABdAtYAXAHNAJgCvABGAv4AXgLtAFMC9wBbAuIAXQKmAKQDEwBZAtAAXQGc/n0GXQB1BnQAdQeUAEYFxQB1BrIAXgaHAHUHdABeB2sAWwa9AKQBvAAlAdn/6wHAACYB3P/9BfkAJQHwADsB9QAaA+MAigPj/+oBwwB8An0AlQMTAI8EjABSAqL/xwKXAJ4CWAB8AlgATQJOAIQCW/+7AngAEAKB/7ECVAA/AlP/wALZAHEC2QBxBC0AcAi6AWYD5gARBdX/9AHk/+sDVP/rA00ArgM0AKQB3wCvAdAApANqAE8DZwA4AhsATwIaADkDSgD+AegA/gHmAAAB5gAAAdMAAAQuAG0FNAB1BBAAVATtAGADzQA1BHj/7ANk/v8Ei//PBJT/1gQ7ACoF4QANBFwAKwSU/9YHyAAuBIsAjAV4AQkFeACXBXgApgV4AP8FeACQBXgASQV4AIEFeACDBXgA3wV4AMMFeACDBXgBAwJwANUEdwAXBCwAjgQsAI4EEABtBCwAjgQvAGwELwBsBBwAWwQcAIkELgBDBC0AQwQwAEQETAB1BEQAkAQsAI4DiwA1BFoAUwV7AFUDTP+gBbgAWgU8/9wE4gCOA/P/7gS2AC8EagBDBG0AEQaRAK8JcwCvA6oAewYxAEYFSgAxBDMAkQOsAAwGEABNBAIAdwYQAE0G+ABKA30A+wIRAEoCEQBKA88AHQNpAEQDbP/tCK0AdgeKACgB4wCnAa8ApgAAAHEAAAB0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ZW50BEhiYXILSGNpcmN1bWZsZXgHdW5pMUUyNAZJYnJldmUHdW5pMDIwOAd1bmkxRUNBB3VuaTFFQzgHdW5pMDIwQQdJbWFjcm9uB0lvZ29uZWsGSXRpbGRlC0pjaXJjdW1mbGV4B3VuaTAxMzYGTGFjdXRlBkxjYXJvbgd1bmkwMTNCBExkb3QGTmFjdXRlBk5jYXJvbgd1bmkwMTQ1B3VuaTFFNDQDRW5nBk9icmV2ZQd1bmkxRUQwB3VuaTFFRDgHdW5pMUVEMgd1bmkxRUQ0B3VuaTFFRDYHdW5pMDIwQwd1bmkxRUNDB3VuaTFFQ0UFT2hvcm4HdW5pMUVEQQd1bmkxRUUyB3VuaTFFREMHdW5pMUVERQd1bmkxRUUwDU9odW5nYXJ1bWxhdXQHdW5pMDIwRQdPbWFjcm9uB3VuaTAxRUELT3NsYXNoYWN1dGUGUmFjdXRlBlJjYXJvbgd1bmkwMTU2B3VuaTAyMTAHdW5pMUU1QQd1bmkwMjEyBlNhY3V0ZQtTY2lyY3VtZmxleAd1bmkwMjE4B3VuaTFFNjIHdW5pMDE4RgRUYmFyBlRjYXJvbgd1bmkwMTYyB3VuaTAyMUEHdW5pMUU2QwZVYnJldmUHdW5pMDIxNAd1bmkxRUU0B3VuaTFFRTYFVWhvcm4HdW5pMUVFOAd1bmkxRUYwB3VuaTFFRUEHdW5pMUVFQwd1bmkxRUVFDVVodW5nYXJ1bWxhdXQHdW5pMDIxNgdVbWFjcm9uB1VvZ29uZWsFVXJpbmcGVXRpbGRlBldhY3V0ZQtXY2lyY3VtZmxleAlXZGllcmVzaXMGV2dyYXZlC1ljaXJjdW1mbGV4B3VuaTFFRjQGWWdyYXZlB3VuaTFFRjYHdW5pMUVGOAZaYWN1dGUKWmRvdGFjY2VudAd1bmkxRTkyBmFicmV2ZQd1bmkxRUFGB3VuaTFFQjcHdW5pMUVCMQd1bmkxRUIzB3VuaTFFQjUHdW5pMUVBNQd1bmkxRUFEB3VuaTFFQTcHdW5pMUVBOQd1bmkxRUFCB3VuaTAyMDEHdW5pMUVBMQd1bmkxRUEzB3VuaTAyMDMHYW1hY3Jvbgdhb2dvbmVrCmFyaW5nYWN1dGUHYWVhY3V0ZQtjY2lyY3VtZmxleApjZG90YWNjZW50BmRjYXJvbgd1bmkxRTBEBmVicmV2ZQZlY2Fyb24HdW5pMUVCRgd1bmkxRUM3B3VuaTFFQzEHdW5pMUVDMwd1bmkxRUM1B3VuaTAyMDUKZWRvdGFjY2VudAd1bmkxRUI5B3VuaTFFQkIHdW5pMDIwNwdlbWFjcm9uB2VvZ29uZWsHdW5pMUVCRAd1bmkwMjU5C2djaXJjdW1mbGV4B3VuaTAxMjMKZ2RvdGFjY2VudARoYmFyC2hjaXJjdW1mbGV4B3VuaTFFMjUGaWJyZXZlB3VuaTAyMDkJaS5sb2NsVFJLB3VuaTFFQ0IHdW5pMUVDOQd1bmkwMjBCB2ltYWNyb24HaW9nb25lawZpdGlsZGUHdW5pMDIzNwtqY2lyY3VtZmxleAd1bmkwMTM3DGtncmVlbmxhbmRpYwZsYWN1dGUGbGNhcm9uB3VuaTAxM0MEbGRvdAZuYWN1dGUGbmNhcm9uB3VuaTAxNDYHdW5pMUU0NQNlbmcGb2JyZXZlB3VuaTFFRDEHdW5pMUVEOQd1bmkxRUQzB3VuaTFFRDUHdW5pMUVENwd1bmkwMjBEB3VuaTFFQ0QHdW5pMUVDRgVvaG9ybgd1bmkxRURCB3VuaTFFRTMHdW5pMUVERAd1bmkxRURGB3VuaTFFRTENb2h1bmdhcnVtbGF1dAd1bmkwMjBGB29tYWNyb24HdW5pMDFFQgtvc2xhc2hhY3V0ZQZyYWN1dGUGcmNhcm9uB3VuaTAxNTcHdW5pMDIxMQd1bmkxRTVCB3VuaTAyMTMGc2FjdXRlC3NjaXJjdW1mbGV4B3VuaTAyMTkHdW5pMUU2MwR0YmFyBnRjYXJvbgd1bmkwMTYzB3VuaTAyMUIHdW5pMUU2RAZ1YnJldmUHdW5pMDFENAd1bmkwMjE1B3VuaTFFRTUHdW5pMUVFNwV1aG9ybgd1bmkxRUU5B3VuaTFFRjEHdW5pMUVFQgd1bmkxRUVEB3VuaTFFRUYNdWh1bmdhcnVtbGF1dAd1bmkwMjE3B3VtYWNyb24HdW9nb25lawV1cmluZwZ1dGlsZGUGd2FjdXRlC3djaXJjdW1mbGV4CXdkaWVyZXNpcwZ3Z3JhdmULeWNpcmN1bWZsZXgHdW5pMUVGNQZ5Z3JhdmUHdW5pMUVGNwd1bmkxRUY5BnphY3V0ZQp6ZG90YWNjZW50B3VuaTFFOTMHemVyby50ZgZvbmUudGYGdHdvLnRmCHRocmVlLnRmB2ZvdXIudGYHZml2ZS50ZgZzaXgudGYIc2V2ZW4udGYIZWlnaHQudGYHbmluZS50Zgl6ZXJvLmRub20Ib25lLmRub20IdHdvLmRub20KdGhyZWUuZG5vbQlmb3VyLmRub20JZml2ZS5kbm9tCHNpeC5kbm9tCnNldmVuLmRub20KZWlnaHQuZG5vbQluaW5lLmRub20JemVyby5udW1yCG9uZS5udW1yCHR3by5udW1yCnRocmVlLm51bXIJZm91ci5udW1yCWZpdmUubnVtcghzaXgubnVtcgpzZXZlbi5udW1yCmVpZ2h0Lm51bXIJbmluZS5udW1yB3VuaTIwNzAHdW5pMDBCOQd1bmkwMEIyB3VuaTAwQjMHdW5pMjA3NAd1bmkyMDc1B3VuaTIwNzYHdW5pMjA3Nwd1bmkyMDc4B3VuaTIwNzkHdW5pMjE1Mwd1bmkyMTU0CW9uZWVpZ2h0aAx0aHJlZWVpZ2h0aHMLZml2ZWVpZ2h0aHMMc2V2ZW5laWdodGhzCXBlcmlvZC50Zghjb21tYS50Zgd1bmkwMEFECmVtZGFzaC5hbHQHdW5pMDBBMAJDUg1jb2xvbm1vbmV0YXJ5BGRvbmcERXVybwRsaXJhB3VuaTIwQkEHdW5pMjBBNgd1bmkyMEI5B3VuaTIwQTkHY2VudC50ZhBjb2xvbm1vbmV0YXJ5LnRmCWRvbGxhci50Zgdkb25nLnRmB0V1cm8udGYJZmxvcmluLnRmCGZyYW5jLnRmB2xpcmEudGYKdW5pMjBCQS50Zgp1bmkyMEI5LnRmC3N0ZXJsaW5nLnRmBnllbi50Zgd1bmkyMjE5B3VuaTIyMTUIZW1wdHlzZXQHdW5pMjEyNgd1bmkyMjA2B3VuaTAwQjUHdW5pMjExNwd1bmkyMTEzB3VuaTIxMTYHdW5pMjEyMAd1bmkwMkJDB3VuaTAyQkIHdW5pMDMwOAd1bmkwMzA3CWdyYXZlY29tYglhY3V0ZWNvbWIHdW5pMDMwQgt1bmkwMzBDLmFsdAd1bmkwMzAyB3VuaTAzMEMHdW5pMDMwNgd1bmkwMzBBCXRpbGRlY29tYgd1bmkwMzA0DWhvb2thYm92ZWNvbWIHdW5pMDMwRgd1bmkwMzExB3VuaTAzMTIHdW5pMDMxQgxkb3RiZWxvd2NvbWIHdW5pMDMyNgd1bmkwMzI3B3VuaTAzMjgQZ3JhdmVjb21iLm5hcnJvdxBhY3V0ZWNvbWIubmFycm93C3VuaTAzMDYwMzAxC3VuaTAzMDYwMzAwC3VuaTAzMDYwMzA5C3VuaTAzMDYwMzAzC3VuaTAzMDIwMzAxC3VuaTAzMDIwMzAwC3VuaTAzMDIwMzA5C3VuaTAzMDIwMzAzB3VuaTAwMDAAAAAAAQAAAAoAKABQAAJERkxUAA5sYXRuAA4ABAAAAAD//wADAAAAAQACAANrZXJuACJtYXJrABpta21rABQAAAABAAMAAAACAAEAAgAAAAEAAAAEEYIQ8AAkAAoABgAQAAEACgAAAAERaBFoAAERXAAMAAERVgAEAAAAAQAIAAEQrg8SAAQPrAAMAXwPAAAADvoO9A8AAAAO+g70DwAAAA76DvQPAAAADvoO9A8AAAAO+g70DwAAAA76DvQPAAAADvoO9A8AAAAO+g70DwAAAA76DvQPAAAADvoO9A8AAAAO+g70DwAAAA76DvQPAAAADvoO9A8AAAAO+g70DwAAAA76DvQPAAAADvoO9A8AAAAO+g70DwAAAA76DvQPAAAADvoO9A8AAAAO+g70DwAAAA76DvQPAAAADvoO9A8AAAAO+g70DwAAAA76DvQPAAAADvoO9A7uAAAAAAAADu4AAAAAAAAO6AAADuIAAA7oAAAO4gAADugAAA7iAAAO6AAADuIAAA7oAAAO4gAADugAAA7iAAAO3AAADtYAAA7QAAAOygAADtwAAA7WAAAO0AAADsoAAA7cAAAO1gAADsQAAA6+DrgOxAAADr4OuA7EAAAOvg64DsQAAA6+DrgOxAAADr4OuA7EAAAOvg64DsQAAA6+DrgOxAAADr4OuA7EAAAOvg64DsQAAA6+DrgOxAAADr4OuA7EAAAOvg64DsQAAA6+DrgOxAAADr4OuA7EAAAOvg64DsQAAA6+DrgOxAAADr4OuA7EAAAOvg64DsQAAA6+DrgOxAAADr4OuA7oAAAOsgAADugAAA6yAAAO6AAADrIAAA7oAAAOsgAADugAAA6yAAAOrAAADqYAAA6gAAAOmgAADqwAAA6mAAAOrAAADqYAAA6UAAAOjg6IDpQAAA6ODogOlAAADo4OiA6UAAAOjg6IDpQAAA6ODogOlAAADo4OiA6UAAAOjg6IDpQAAA6ODogOlAAADo4OiA6UAAAOjg6IDpQAAA6ODogOlAAADo4OiA6UAAAOjg6IDpQAAA6ODogOggAAAAAAAA6CAAAAAAAADnwAAA52AAAOfAAADnYAAA5wAAAOagAADnAAAA5qAAAOcAAADmoAAA5wAAAOagAADnAAAA5qAAAOZAAADr4AAA5eAAAOWAAADl4AAA5YAAAOXgAADlgAAA5eAAAOWAAADl4AAA5YAAAOXgAADlgAAA5eAAAOWAAADlIOTA5GDkAOUg5MDkYOQA5SDkwORg5ADlIOTA5GDkAOUg5MDkYOQA5SDkwORg5ADlIOTA5GDkAOUg5MDkYOQA5SDkwORg5ADlIOTA5GDkAOUg5MDkYOQA5SDkwORg5ADlIOTA5GDkAOUg5MDkYOQA5SDkwORg5ADlIOTA5GDkAOUg5MDkYOQA5SDkwORg5ADlIOTA5GDkAOUg5MDkYOQA5SDkwORg5ADlIOTA5GDkAOUg5MDkYOQA5SDkwORg5ADjoAAAAAAAAOOgAAAAAAAA5SDkwORg5ADjQAAAAAAAAOLgAADigAAA4uAAAOKAAADi4AAA4oAAAOLgAADigAAA4uAAAOKAAADi4AAA4oAAAOLgAADigAAA4iAAAOHAAADiIAAA4cAAAOIgAADhwAAA4iAAAOHAAADiIAAA4cAAAOIgAADhwAAA4iAAAOHAAADhYAAA4QAAAOFgAADhAAAA4WAAAOEAAADhYAAA4QAAAOFgAADhAAAA4WAAAOEAAADgoOBA3+DfgOCg4EDf4N+A4KDgQN/g34DgoOBA3+DfgOCg4EDf4N+A4KDgQN/g34DgoOBA3+DfgOCg4EDf4N+A4KDgQN/g34DgoOBA3+DfIOCg4EDf4N+A4KDgQN/g34DgoOBA3+DfgOCg4EDf4N+A4KDgQN/g34DgoOBA3+DfgOCg4EDf4N+A4KDgQN/g34DgoOBA3+DfgOCg4EDf4N+A4KDgQN/g34DewAAAAAAAAN7AAAAAAAAA3sAAAAAAAADewAAAAAAAAN7AAAAAAAAA3mAAAN4AAADeYAAA3gAAAN5gAADeAAAA3mAAAN4AAADeYAAA3gAAAN5gAADeAAAA3mAAAN4AAADeYAAA3gAAAN2gAADdQAAA3aAAAN1AAADdoAAA3UAAAN2gAADdQAAA3aAAAN1AAADc4AAA3IDcINzgAADcgNwg3OAAANyA3CDc4AAA3IDcINzgAADcgNwg3OAAANyA3CDc4AAA3IDcINzgAADcgNwg3OAAANyA3CDc4AAA3IDcINzgAADcgNwg3OAAANyA3CDc4AAA3IDcINzgAADcgNwg28AAANtg2wDc4AAA3IDcINzgAADcgNwg2qAAANpA2eDc4AAA3IDcINzgAADcgNwg3OAAANyA3CDc4AAA3IDcINzgAADcgNwg3OAAANyA3CDc4AAA3IDcINmAAADZINjA2YAAANkg2MDYYAAA2AAAANhgAADYAAAA2GAAANgAAADYYAAA2AAAANhgAADYAAAA2GAAANgAAAAAAAAA16AAAAAAAADXoAAAAAAAANegAAAAAAAA16AAANdAAADW4NaA10AAANbg1oDXQAAA1uDWgNdAAADW4NaA10AAANbg1oDXQAAA1uDWgNdAAADW4NaA10AAANbg1oDXQAAA1uDWgNdAAADW4NaA10AAANbg1oDXQAAA1uDWgNdAAADW4NaA10AAANbg1oDXQAAA1uDWgNdAAADW4NaA10AAANbg1oDXQAAA1uDWgNdAAADW4NaA10AAANbg1oDWIAAA1cDVYNUAAAAAAAAA1QAAAAAAAADVAAAAAAAAANUAAAAAAAAA1QAAAAAAAADUoAAA1EAAANPgAADTgAAA1KAAANRAAADUoAAA1EAAAAAAAADTINLA0mAAAAAA0sDSYAAAAADSwNJgAAAAANLA0mAAAAAA0sDSYAAAAADSwNJgAAAAANLA0mAAAAAA0sAAAAAA0yDSwNJgAAAAANLA0mAAAAAA0sDSYAAAAADSwNJgAAAAANLAAAAAANMg0sDSYAAAAADSwNIAAAAAAAAA0gAAAAAAAAAAAAAA0aAAAAAAAADRoAAA0UAAANDgAADRQAAA0OAAANFAAADQ4AAA0UAAANDgAADRQAAA0OAAANCAAADQIAAAz8AAAM9gAADPwAAAz2AAAM/AAADPYAAAz8AAAM9gAADPwAAAz2AAAM/AAADPYAAAzwDOoM5AzeDPAM6gzkDN4M8AzqDOQM3gzwDOoM5AzeDPAM6gzkDN4M8AzqDOQM3gzwDOoM5AzeDPAM6gzkDN4M8AzqDOQM3gzwDOoM5AzeDPAM6gzkDN4M8AzqDOQM3gzwDOoM5AzeDPAM6gzkDN4M8AzqDOQM3gzwDOoM5AzeDPAM6gzkDN4M2AzSDRoMzAzw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';

    function novoDocPDF() {
        const ctor = (window.jspdf && window.jspdf.jsPDF) || window.jsPDF;
        if (!ctor) throw new Error('biblioteca jsPDF não carregada');
        const doc = new ctor({ orientation: 'portrait', unit: 'mm', format: 'a4' });
        if (typeof doc.autoTable !== 'function') throw new Error('plugin autoTable não carregado');
        doc.addFileToVFS('PublicSans-Regular.ttf', FONTE_PUBLIC_SANS_REGULAR);
        doc.addFont('PublicSans-Regular.ttf', 'PublicSans', 'normal');
        doc.addFileToVFS('PublicSans-Bold.ttf', FONTE_PUBLIC_SANS_BOLD);
        doc.addFont('PublicSans-Bold.ttf', 'PublicSans', 'bold');
        doc.addFileToVFS('PublicSans-Italic.ttf', FONTE_PUBLIC_SANS_ITALIC);
        doc.addFont('PublicSans-Italic.ttf', 'PublicSans', 'italic');
        doc.setFont('PublicSans', 'normal');
        return doc;
    }

    // somenteResumo: quando true, gera só a página de resumo (KPIs + gráficos), sem a
    // tabela discriminada de processos.
    function gerarPDF(dados, cfg, somenteResumo) {
        const doc = novoDocPDF();
        montarResumoGenerico(doc, dados, cfg, true, false);
        doc.outline.add(null, 'Resumo', { pageNumber: 1 });
        if (!somenteResumo) {
            const pgTabela = montarTabelaGenerico(doc, dados, cfg, false);
            doc.outline.add(null, 'Tabela detalhada', { pageNumber: pgTabela });
        }
        const sufixo = somenteResumo ? '_resumo' : '';
        baixarBlob(doc.output('blob'), `${cfg.nomeArquivo}${sufixo}_${dataArquivo()}.pdf`);
    }

    // ── PDF de Conclusões — uma seção por juiz responsável ──────────────────────
    // Pensado para "expedir" individualmente: cada juiz ganha um resumo (quantas
    // pendentes, por tipo, por agrupador, quantas com/sem pré-análise) seguido da
    // tabela discriminada dos processos dele, com marcador (bookmark) próprio.

    // "Sim"/"Não" a partir do texto bruto da célula de Pré-análise (não vazio = houve).
    function temPreAnalise(d) { return !!(d.preAnalise && d.preAnalise.trim()); }

    // Dias entre hoje e a data extraída da pré-análise (preAnaliseData); '' quando essa
    // data não pôde ser identificada no texto da célula (ver comentário em CFG_CONCLUSOES).
    function diasDesdePreAnalise(d, now) {
        if (!d.preAnaliseData) return '';
        return diasDecorridos(d.preAnaliseData, now);
    }

    function montarResumoJuizConclusoes(doc, juiz, sub, now, primeira) {
        if (!primeira) doc.addPage();
        const agora = new Date();
        const pw = doc.internal.pageSize.getWidth();
        const ph = doc.internal.pageSize.getHeight();
        const m = 12;
        const uw = pw - 2 * m;
        const hoje = agora.toLocaleDateString('pt-BR');
        const hora = agora.toLocaleTimeString('pt-BR', { hour: '2-digit', minute: '2-digit' });
        const gap = 6;

        const comPreAnalise = sub.filter(temPreAnalise);
        const semPreAnalise = sub.length - comPreAnalise.length;
        const prio = contarPrioritarios(sub);

        let hy = m + 2;
        doc.setFont('PublicSans', 'bold'); doc.setFontSize(16); doc.setTextColor(...COR.tinta);
        doc.text(TITULO_CONCLUSOES_POR_JUIZ, m, hy);
        hy += 8;
        doc.setFont('PublicSans', 'bold'); doc.setFontSize(11.5); doc.setTextColor(...COR.azul);
        const linhasJuiz = doc.splitTextToSize('Juiz(a): ' + juiz, uw);
        doc.text(linhasJuiz, m, hy);
        hy += linhasJuiz.length * 5.2 + 1.5;
        doc.setFont('PublicSans', 'normal'); doc.setFontSize(9); doc.setTextColor(...COR.tintaSec);
        doc.text(`Extraído em ${hoje} às ${hora}  •  ${sub.length} conclusão(ões) pendente(s)`, m, hy);
        hy += 3;
        doc.setDrawColor(...COR.azul); doc.setLineWidth(0.5); doc.line(m, hy, pw - m, hy);

        const kY = hy + 6;
        const kpis = [
            { titulo: 'Pendentes', valor: String(sub.length), subs: [], acento: COR.azul },
            { titulo: 'Prioritários', valor: String(prio), subs: [`${sub.length ? Math.round(prio / sub.length * 100) : 0}% do total`], acento: COR.vermelho },
            { titulo: 'Com pré-análise', valor: String(comPreAnalise.length), subs: [`${sub.length ? Math.round(comPreAnalise.length / sub.length * 100) : 0}% do total`], acento: COR.aqua },
            { titulo: 'Sem pré-análise', valor: String(semPreAnalise), subs: [], acento: COR.ambar },
        ];
        const kW = (uw - (kpis.length - 1) * gap) / kpis.length;
        kpis.forEach((k, i) => desenharCard(doc, m + i * (kW + gap), kY, kW, 28, k.titulo, k.valor, k.subs, true, k.acento));

        const gY0 = kY + 28 + gap + 2;
        const charts = [
            { tipo: 'barras', span: 1, titulo: 'Pendentes por Tipo de Conclusão', itens: contarPorCampo(sub, 'tipoConclusao', 10) },
            { tipo: 'barras', span: 1, titulo: 'Pendentes por Agrupador', itens: contarPorCampo(sub, 'agrupador', 10) },
        ];
        desenharGradeGraficos(doc, m, gY0, uw, ph - m - gY0, charts);
        desenharRodape(doc, TITULO_CONCLUSOES_POR_JUIZ, `${hoje} ${hora}`, pw, ph, m, false);
    }

    function montarTabelaJuizConclusoes(doc, juiz, sub, now, link) {
        const agora = new Date();
        const pw = doc.internal.pageSize.getWidth();
        const ph = doc.internal.pageSize.getHeight();
        const m = 12;
        const hoje = agora.toLocaleDateString('pt-BR');
        const hora = agora.toLocaleTimeString('pt-BR', { hour: '2-digit', minute: '2-digit' });

        // Prioritárias primeiro; dentro de cada grupo, da data de conclusão mais antiga
        // para a mais nova.
        const ordenados = sub.slice().sort((a, b) => {
            const pa = a.prioritario ? 0 : 1, pb = b.prioritario ? 0 : 1;
            if (pa !== pb) return pa - pb;
            const ta = parseDataBR(a.dtRemessa); const tb = parseDataBR(b.dtRemessa);
            return (ta == null ? Infinity : ta) - (tb == null ? Infinity : tb);
        });

        doc.addPage();
        const paginaInicial = doc.internal.getNumberOfPages();
        tituloSecao(doc, m, m + 3, pw - 2 * m, `Tabela discriminada — ${juiz}`);
        const tabInicioY = m + 8;

        doc.autoTable({
            // As duas últimas colunas são sempre Dt. Conclusão e Dias desde pré-análise.
            columns: [
                { header: 'Processo', dataKey: 'processo' },
                { header: 'Classe', dataKey: 'classe' },
                { header: 'Tipo de Conclusão', dataKey: 'tipo' },
                { header: 'Agrupador', dataKey: 'agrupador' },
                { header: 'Pré-análise', dataKey: 'preAnalise' },
                { header: 'Dias', dataKey: 'dias' },
                { header: 'Dt. Conclusão', dataKey: 'dtRemessa' },
                { header: 'Dias desde pré-análise', dataKey: 'diasPreAnalise' },
            ],
            body: ordenados.map(d => ({
                processo: d.processo, classe: d.classe, tipo: d.tipoConclusao, agrupador: d.agrupador,
                preAnalise: temPreAnalise(d) ? 'Sim' : 'Não',
                dias: diasDecorridos(d.dtRemessa, now),
                dtRemessa: d.dtRemessa,
                diasPreAnalise: diasDesdePreAnalise(d, now) || '—',
            })),
            startY: tabInicioY,
            margin: { left: m, right: m, top: m, bottom: 14 },
            theme: 'grid',
            styles: { font: 'PublicSans', fontSize: 7.5, cellPadding: 1.6, textColor: COR.tintaSec,
                      lineColor: COR.grade, lineWidth: 0.1, overflow: 'linebreak', valign: 'middle' },
            headStyles: { fillColor: COR.azul, textColor: [255, 255, 255], fontStyle: 'bold', fontSize: 8 },
            alternateRowStyles: { fillColor: COR.cartao },
            columnStyles: {
                processo: { cellWidth: 28 }, classe: { cellWidth: 20 }, tipo: { cellWidth: 26 },
                agrupador: { cellWidth: 22 }, preAnalise: { cellWidth: 16 }, dias: { cellWidth: 12 },
                dtRemessa: { cellWidth: 18 }, diasPreAnalise: { cellWidth: 20 },
            },
            didParseCell: (data) => {
                if (data.section === 'body' && data.column.dataKey === 'processo' && ordenados[data.row.index] && ordenados[data.row.index].prioritario) {
                    data.cell.styles.textColor = COR_PRIORITARIO;
                    data.cell.styles.fontStyle = 'bold';
                }
            },
            didDrawPage: () => desenharRodape(doc, `Tabela — ${juiz}`, `${hoje} ${hora}`, pw, ph, m, link),
        });

        return paginaInicial;
    }

    const TITULO_CONCLUSOES_POR_JUIZ = 'Conclusões Pendentes por Juiz';

    function gerarPDFConclusoesPorJuiz(dados) {
        const doc = novoDocPDF();
        const now = Date.now();
        const juizes = [...new Set(dados.map(d => (d.responsavel || '').trim() || '(sem responsável)'))]
            .sort((a, b) => a.localeCompare(b, 'pt-BR'));
        juizes.forEach((juiz, i) => {
            const sub = dados.filter(d => ((d.responsavel || '').trim() || '(sem responsável)') === juiz);
            const pgResumo = doc.internal.getNumberOfPages() + (i === 0 ? 0 : 1);
            montarResumoJuizConclusoes(doc, juiz, sub, now, i === 0);
            const bmJuiz = doc.outline.add(null, `${juiz} (${sub.length})`, { pageNumber: pgResumo });
            const pgTabela = montarTabelaJuizConclusoes(doc, juiz, sub, now);
            doc.outline.add(bmJuiz, 'Tabela detalhada', { pageNumber: pgTabela });
        });
        baixarBlob(doc.output('blob'), `conclusoes_por_juiz_projudi_${dataArquivo()}.pdf`);
    }

    // Resolve como montar o resumo/tabela de uma seção do PDF conjunto, dado o cfg do
    // relatório (genérico via cfg.pdf, ou o caso especial do Tempo Médio).
    function descreverSecaoPDF(cfg, somenteResumo) {
        if (cfg === CFG_TEMPOMEDIO) {
            return {
                rotulo: 'Tempo Médio de Cumprimento',
                montarResumo: (doc, dados, primeira, comIndice) => montarResumoTempoMedio(doc, dados, primeira, comIndice),
                montarTabela: (doc, dados, comIndice) => montarTabelaTempoMedio(doc, dados, comIndice),
            };
        }
        if (cfg === CFG_PARALISADOS) {
            return {
                rotulo: TITULO_PARALISADOS,
                montarResumo: (doc, dados, primeira, comIndice) => montarResumoParalisados(doc, dados, primeira, comIndice),
                montarTabela: (doc, dados, comIndice) => montarTabelaParalisados(doc, dados, comIndice),
            };
        }
        if (cfg === CFG_REMESSAS) {
            return {
                rotulo: TITULO_REMESSAS,
                montarResumo: (doc, dados, primeira, comIndice) => montarResumoRemessas(doc, dados, primeira, comIndice),
                montarTabela: (doc, dados, comIndice) => montarTabelaRemessas(doc, dados, comIndice),
            };
        }
        if (cfg === CFG_AUDIENCIAS) {
            return {
                rotulo: TITULO_AUDIENCIAS,
                // Resumo e tabela sempre juntos (pedido do usuário) — a tabela é desenhada
                // dentro do próprio montarResumoAudiencias, então não há passo de tabela
                // separado aqui (ver secaoTemTabela em gerarPDFConjunto).
                montarResumo: (doc, dados, primeira, comIndice) => montarResumoAudiencias(doc, dados, primeira, comIndice, somenteResumo),
                montarTabela: null,
            };
        }
        if (cfg === CFG_AUDIENCIAS_DESIGNADAS) {
            return {
                rotulo: TITULO_AUDIENCIAS_DESIGNADAS,
                montarResumo: (doc, dados, primeira, comIndice) => montarResumoAudienciasDesignadas(doc, dados && dados[0], primeira, comIndice),
                montarTabela: (doc, dados, comIndice) => montarTabelaAudienciasDesignadas(doc, (dados && dados[0] && dados[0].tabela) || [], comIndice),
            };
        }
        if (cfg === CFG_AUDIENCIAS_REALIZADAS) {
            return {
                rotulo: TITULO_AUDIENCIAS_REALIZADAS,
                // Só resumo — não há lista de processos a discriminar, apenas totais (geral
                // e por usuário) — ver secaoTemTabela em gerarPDFConjunto.
                montarResumo: (doc, dados, primeira, comIndice) => montarResumoAudienciasRealizadas(doc, dados && dados[0], primeira, comIndice),
                montarTabela: null,
            };
        }
        return {
            rotulo: cfg.pdf.titulo,
            montarResumo: (doc, dados, primeira, comIndice) => montarResumoGenerico(doc, dados, cfg, primeira, comIndice),
            montarTabela: (doc, dados, comIndice) => montarTabelaGenerico(doc, dados, cfg, comIndice),
        };
    }

    // Desenha uma caixa de domínio (Cartório ou Gabinete) na página "Situação da
    // Unidade": cabeçalho, KPIs e uma mini-tabela por item (tarefa do Cartório, ou
    // magistrado(a) do Gabinete) — a situação (Crítico/Atenção/Regular) só aparece por
    // item, na última coluna da tabela, nunca como veredito agregado do domínio inteiro.
    // Retorna o Y logo abaixo da tabela.
    function desenharBlocoDominio(doc, x, y, w, cfg) {
        const pw = doc.internal.pageSize.getWidth();

        const headH = 15;
        doc.setDrawColor(...COR.grade); doc.setLineWidth(0.3); doc.setFillColor(...COR.cartao);
        doc.rect(x, y, w, headH, 'FD');
        doc.setFont('PublicSans', 'bold'); doc.setFontSize(13); doc.setTextColor(...COR.tinta);
        doc.text(cfg.titulo, x + 6, y + 9.5);

        let yy = y + headH + 5;
        // Observação (substitui o antigo subtítulo com a lista de tarefas) — explica o
        // critério de situação usado especificamente neste domínio.
        if (cfg.observacao) {
            doc.setFont('PublicSans', 'italic'); doc.setFontSize(7.8); doc.setTextColor(...COR.muted);
            const linhasObs = doc.splitTextToSize(cfg.observacao, w - 12);
            doc.text(linhasObs, x + 6, yy);
            yy += linhasObs.length * 3.4 + 5;
        }

        if (!cfg.itens.length) return yy;
        // colunaExtra (opcional): coluna adicional entre Prioritários e Mais antiga,
        // calculada a partir dos dados brutos do item (ex.: pré-analisados no Gabinete).
        const temExtra = !!cfg.colunaExtra;
        const corpo = cfg.itens.map(it => {
            const infoIt = SITUACAO_INFO[it.status] || SITUACAO_INFO.regular;
            const linha = {
                rotulo: it.rotulo, pendentes: String(it.pendentes), prioritarios: String(it.prioritarios),
                antiga: it.maisAntiga != null ? `${it.maisAntiga} dias` : '—', situacao: infoIt.rotulo, _cor: infoIt.cor,
            };
            if (temExtra) linha.extra = String(cfg.colunaExtra.get(it.dados || []));
            return linha;
        });
        // Linha de total — substitui os KPIs de agregado que existiam acima da tabela (o
        // resumo inicial agora fica só em tabela, sem cards).
        const totalPendentes = cfg.itens.reduce((s, it) => s + it.pendentes, 0);
        const totalPrioritarios = cfg.itens.reduce((s, it) => s + it.prioritarios, 0);
        const totalExtra = temExtra ? cfg.itens.reduce((s, it) => s + (cfg.colunaExtra.get(it.dados || []) || 0), 0) : 0;
        const maisAntigaGeral = cfg.itens.reduce((m, it) => (it.maisAntiga != null && (m == null || it.maisAntiga > m) ? it.maisAntiga : m), null);
        corpo.push({
            rotulo: 'Total', pendentes: String(totalPendentes), prioritarios: String(totalPrioritarios),
            extra: temExtra ? String(totalExtra) : undefined,
            antiga: maisAntigaGeral != null ? `${maisAntigaGeral} dias` : '—', situacao: '', _cor: null, _total: true,
        });
        const colunas = [
            { header: cfg.colunaRotulo, dataKey: 'rotulo' },
            { header: 'Pendentes', dataKey: 'pendentes' },
            { header: 'Prioritários', dataKey: 'prioritarios' },
        ];
        if (temExtra) colunas.push({ header: cfg.colunaExtra.header, dataKey: 'extra' });
        colunas.push({ header: 'Mais antiga', dataKey: 'antiga' }, { header: 'Situação', dataKey: 'situacao' });
        const columnStyles = {
            rotulo: { cellWidth: w * (temExtra ? 0.28 : 0.34), fontStyle: 'bold', textColor: COR.tinta },
            pendentes: { cellWidth: w * (temExtra ? 0.13 : 0.16), halign: 'right' },
            prioritarios: { cellWidth: w * (temExtra ? 0.13 : 0.16), halign: 'right' },
            antiga: { cellWidth: w * (temExtra ? 0.14 : 0.16), halign: 'right' },
            situacao: { cellWidth: w * (temExtra ? 0.16 : 0.18) },
        };
        if (temExtra) columnStyles.extra = { cellWidth: w * 0.16, halign: 'right' };
        doc.autoTable({
            columns: colunas,
            body: corpo,
            startY: yy,
            margin: { left: x, right: pw - x - w, bottom: 14 },
            theme: 'grid',
            styles: { font: 'PublicSans', fontSize: 8.5, cellPadding: 2.2, textColor: COR.tintaSec, lineColor: COR.grade, lineWidth: 0.1, valign: 'middle' },
            headStyles: { fillColor: COR.azul, textColor: [255, 255, 255], fontStyle: 'bold', fontSize: 8 },
            alternateRowStyles: { fillColor: COR.cartao },
            columnStyles,
            didParseCell: (data) => {
                if (data.section === 'body' && corpo[data.row.index]._total) {
                    data.cell.styles.fontStyle = 'bold';
                    data.cell.styles.fillColor = COR.cartao;
                }
                if (data.section === 'body' && data.column.dataKey === 'situacao' && !corpo[data.row.index]._total) {
                    data.cell.styles.textColor = corpo[data.row.index]._cor;
                    data.cell.styles.fontStyle = 'bold';
                }
            },
        });
        return doc.lastAutoTable.finalY;
    }

    // Bloco do Cartório na "Situação da Unidade": mesmo cabeçalho/KPIs de
    // desenharBlocoDominio, mas com uma tabela de colunas GENÉRICAS (Tarefa/Indicador/
    // Detalhamento/Situação) em vez de Pendentes/Prioritários/Mais antiga — precisa disso
    // porque agora TODOS os relatórios entram como uma linha do Cartório (pedido do
    // usuário), inclusive Tempo Médio e as três Audiências, que não têm o mesmo formato de
    // "pendentes" das tarefas clássicas (Juntadas, Retorno, Paralisados, Remessas,
    // Suspensos). Gabinete continua usando desenharBlocoDominio, sem mudança.
    function desenharBlocoCartorioUnificado(doc, x, y, w, cfg) {
        const pw = doc.internal.pageSize.getWidth();

        const headH = 15;
        doc.setDrawColor(...COR.grade); doc.setLineWidth(0.3); doc.setFillColor(...COR.cartao);
        doc.rect(x, y, w, headH, 'FD');
        doc.setFont('PublicSans', 'bold'); doc.setFontSize(13); doc.setTextColor(...COR.tinta);
        doc.text(cfg.titulo, x + 6, y + 9.5);

        let yy = y + headH + 5;
        if (cfg.observacao) {
            doc.setFont('PublicSans', 'italic'); doc.setFontSize(7.8); doc.setTextColor(...COR.muted);
            const linhasObs = doc.splitTextToSize(cfg.observacao, w - 12);
            doc.text(linhasObs, x + 6, yy);
            yy += linhasObs.length * 3.4 + 5;
        }

        if (!cfg.linhas.length) return yy;

        doc.autoTable({
            columns: [
                { header: 'Tarefa', dataKey: 'nome' },
                { header: 'Indicador', dataKey: 'indicador' },
                { header: 'Detalhamento', dataKey: 'detalhamento' },
                { header: 'Situação', dataKey: 'situacao' },
            ],
            body: cfg.linhas.map(l => ({ nome: l.nome, indicador: l.indicador, detalhamento: l.detalhamento, situacao: l.semSituacao ? '—' : l.situacaoLabel })),
            startY: yy,
            margin: { left: x, right: pw - x - w, bottom: 14 },
            theme: 'grid',
            styles: { font: 'PublicSans', fontSize: 8.5, cellPadding: 2.2, textColor: COR.tintaSec, lineColor: COR.grade, lineWidth: 0.1, valign: 'middle' },
            headStyles: { fillColor: COR.azul, textColor: [255, 255, 255], fontStyle: 'bold', fontSize: 8 },
            alternateRowStyles: { fillColor: COR.cartao },
            columnStyles: {
                nome: { cellWidth: w * 0.26, fontStyle: 'bold', textColor: COR.tinta },
                indicador: { cellWidth: w * 0.19 },
                detalhamento: { cellWidth: w * 0.38 },
                situacao: { cellWidth: w * 0.17, halign: 'right' },
            },
            didParseCell: (data) => {
                if (data.section === 'body' && data.column.dataKey === 'situacao') {
                    const l = cfg.linhas[data.row.index];
                    if (!l.semSituacao) { data.cell.styles.textColor = l.corTexto; data.cell.styles.fontStyle = 'bold'; }
                    else data.cell.styles.textColor = COR.muted;
                }
            },
            // Guarda a posição/página da célula "Tarefa" de cada linha para poder virar um
            // link clicável até a seção detalhada do relatório (ver PASSO 4 em
            // gerarPDFConjunto) — só depois que a página de destino é conhecida.
            didDrawCell: (data) => {
                if (data.section === 'body' && data.column.dataKey === 'nome') {
                    const l = cfg.linhas[data.row.index];
                    if (l) l._rect = { x: data.cell.x, y: data.cell.y, w: data.cell.width, h: data.cell.height, page: doc.internal.getCurrentPageInfo().pageNumber };
                }
            },
        });
        return doc.lastAutoTable.finalY;
    }

    // Página "Situação da Unidade" (substitui a antiga capa/índice): processos ativos
    // por atuação, seguidos de Cartório (tramitação) e Gabinete (decisão) — a situação
    // (Crítico/Atenção/Regular) só aparece por item, nas mini-tabelas de cada domínio.
    function desenharCapaSituacao(doc, cartorio, gabinete, mapaAtivos, agora, primeira) {
        if (!primeira) doc.addPage();
        const pw = doc.internal.pageSize.getWidth();
        const m = 16;
        const uw = pw - 2 * m;
        const hoje = agora.toLocaleDateString('pt-BR');
        const hora = agora.toLocaleTimeString('pt-BR', { hour: '2-digit', minute: '2-digit' });

        doc.setFillColor(...COR.azul); doc.rect(0, 0, pw, 26, 'F');
        doc.setFont('PublicSans', 'bold'); doc.setFontSize(20); doc.setTextColor(255, 255, 255);
        doc.text('Situação da Unidade', m, 17);
        doc.setFont('PublicSans', 'normal'); doc.setFontSize(10); doc.setTextColor(...COR.tintaSec);
        doc.text(`Projudi — TJPR  •  Extraído em ${hoje} às ${hora}`, m, 36);

        let y = 44;

        // Processos Ativos agora é uma LINHA da tabela do Cartório (ver linhasCartorio em
        // gerarPDFConjunto), não mais um card à parte — resumo inicial ficou só em tabela.
        y = desenharBlocoCartorioUnificado(doc, m, y, uw, {
            titulo: 'Cartório',
            observacao: 'Pendências clássicas: situação pela pendência mais antiga (Regular até 30 dias, Atenção '
                + '31–90, Crítico acima de 90). Tempo Médio e Audiências Pendentes/Realizadas são informativos, sem '
                + 'situação. Audiências Designadas segue a régua de 180/360 dias até a audiência mais distante.',
            linhas: cartorio.linhas,
        });

        y += 8;

        desenharBlocoDominio(doc, m, y, uw, {
            titulo: 'Gabinete',
            observacao: 'Situação calculada pela pendência mais antiga de cada magistrado(a). Regular até 30 '
                + 'dias, Atenção de 31 a 120 dias, Crítico acima de 120 dias.',
            situacao: gabinete.situacao,
            colunaRotulo: 'Magistrado(a)',
            itens: gabinete.itens,
            colunaExtra: { header: 'Pré-analisados', get: (dados) => dados.filter(d => !!d.preAnalise).length },
        });
    }

    // PDF único com os relatórios coletados, organizado em duas frentes: CARTÓRIO
    // (Juntadas, Retorno de Conclusos, Paralisados, Remessas Pendentes — tramitação da
    // secretaria) e GABINETE (Conclusões, uma seção por magistrado(a) — trabalho de
    // decisão). Abre com a página "Situação da Unidade" (veredito objetivo de cada
    // frente), seguida das seções do Cartório e depois do Gabinete, cada uma com
    // marcadores (bookmarks) no leitor de PDF. somenteResumo omite as tabelas
    // discriminadas, mantendo só os resumos (KPIs e gráficos) de cada seção.
    // Se a seção "s" (já passada por descreverSecaoPDF) tem algo a discriminar em tabela.
    // A maioria guarda os registros direto em s.dados; Audiências Designadas guarda um
    // resumo único (s.dados = [resumo]) com a lista de linhas dentro de resumo.tabela.
    function secaoTemTabela(s) {
        // Audiências Pendentes desenha a tabela dentro do próprio resumo (pedido do
        // usuário: sempre juntos) — nunca entra no passo de tabela separado.
        if (s.cfgOriginal === CFG_AUDIENCIAS) return false;
        // Audiências Realizadas é só totais (geral + por usuário) — nunca tem tabela
        // discriminada de processos.
        if (s.cfgOriginal === CFG_AUDIENCIAS_REALIZADAS) return false;
        if (s.cfgOriginal === CFG_AUDIENCIAS_DESIGNADAS) {
            const resumo = s.dados && s.dados[0];
            return !!(resumo && resumo.tabela && resumo.tabela.length);
        }
        return s.dados.length > 0;
    }

    function gerarPDFConjunto(secoesEntrada, somenteResumo) {
        const doc = novoDocPDF();
        const agora = new Date();
        const now = agora.getTime();
        const secoes = secoesEntrada.map(s => Object.assign({ dados: s.dados, cfgOriginal: s.cfg }, descreverSecaoPDF(s.cfg, somenteResumo)));

        // Suspensos por Prazo Indeterminado é mais uma tarefa do Cartório (mesmo esquema
        // genérico de Juntadas/Retorno, via cfg.pdf) — não precisa de página própria.
        const CFGS_CARTORIO = [CFG_JUNTADAS, CFG_RETORNO, CFG_PARALISADOS, CFG_REMESSAS, CFG_SUSPENSOS];
        // Seções com cfg.mostrarSeVazio (Suspensos, Audiências Pendentes) aparecem mesmo
        // com dados.length === 0, desde que já tenham sido coletadas (ver KEY_COLETADO/
        // foiColetado) — "zero pendências" é um dado, não um vazio a esconder.
        const secoesCartorio = secoes.filter(s => CFGS_CARTORIO.includes(s.cfgOriginal)
            && (s.dados.length || (s.cfgOriginal.mostrarSeVazio && foiColetado(s.cfgOriginal))));
        const secaoGabinete = secoes.find(s => s.cfgOriginal === CFG_CONCLUSOES && s.dados.length);
        // Seções fora do esquema Cartório/Gabinete (Tempo Médio, Audiências Pendentes)
        // entram depois, sem capa/veredito dedicado — mantém o relatório funcional mesmo
        // nesse caso.
        const CFGS_FORA_DO_ESQUEMA = [...CFGS_CARTORIO, CFG_CONCLUSOES];
        const outrasSecoes = secoes.filter(s => !CFGS_FORA_DO_ESQUEMA.includes(s.cfgOriginal)
            && (s.dados.length || (s.cfgOriginal.mostrarSeVazio && foiColetado(s.cfgOriginal))));

        // Limites de dias (pendência mais antiga) por domínio — ver legenda em
        // desenharCapaSituacao. Cartório: regular ≤30, atenção 31–90, crítico >90.
        // Gabinete: regular ≤30, atenção 31–120, crítico >120.
        const LIMITES_CARTORIO = { atencao: 30, critico: 90 };
        const LIMITES_GABINETE = { atencao: 30, critico: 120 };

        const itensCartorio = secoesCartorio.map(s => {
            const itens = itensParaClassificacao(s.dados, s.cfgOriginal, now);
            const maisAntiga = maiorDias(itens);
            return {
                rotulo: s.rotulo, dados: s.dados, secao: s, _itens: itens,
                pendentes: s.dados.length,
                prioritarios: contarPrioritarios(s.dados),
                maisAntiga,
                status: classificarSituacaoPorDias(maisAntiga, LIMITES_CARTORIO.atencao, LIMITES_CARTORIO.critico),
            };
        });
        // Juntadas/Retorno de Conclusos entram como KPIs próprios no card do Cartório (além
        // da mini-tabela por tarefa, que já os separa) — pedido do usuário. null quando a
        // seção nem chegou a ser coletada (não confundir com "coletado, zero pendências").
        const itemCartorioDe = (cfg) => itensCartorio.find(t => t.secao.cfgOriginal === cfg);
        const itemJuntadas = itemCartorioDe(CFG_JUNTADAS);
        const itemRetorno = itemCartorioDe(CFG_RETORNO);
        // Audiências Pendentes não é uma tarefa do Cartório no esquema (fica em
        // "outrasSecoes"), mas o usuário quer o total também aqui, como KPI — ver
        // "Termos de Audiência Pendentes" abaixo.
        const secaoAudiencias = secoes.find(s => s.cfgOriginal === CFG_AUDIENCIAS);
        const cartorio = {
            itens: itensCartorio,
            totalPendentes: itensCartorio.reduce((s, t) => s + t.pendentes, 0),
            totalPrioritarios: itensCartorio.reduce((s, t) => s + t.prioritarios, 0),
            acima30: itensCartorio.reduce((s, t) => s + t._itens.filter(it => it.dias != null && it.dias > 30).length, 0),
            maisAntiga: maiorDias(itensCartorio.flatMap(t => t._itens)),
            situacao: piorSituacao(itensCartorio.map(t => t.status)),
            juntadasPendentes: itemJuntadas ? itemJuntadas.pendentes : null,
            retornoPendentes: itemRetorno ? itemRetorno.pendentes : null,
            audienciasPendentes: secaoAudiencias ? secaoAudiencias.dados.length : null,
        };

        // Tabela única do Cartório (pedido do usuário): TODOS os relatórios — inclusive
        // Tempo Médio e as três Audiências — como uma linha só, cada um com seu indicador
        // e detalhamento próprios (não têm o mesmo formato de "pendentes/prioritários" das
        // tarefas clássicas). Isso é só o RESUMO na capa — cada um continua tendo sua
        // própria seção com gráfico/tabela discriminada mais adiante (ver outrasSecoes).
        const secaoTempoMedio = secoes.find(s => s.cfgOriginal === CFG_TEMPOMEDIO);
        const secaoAudienciasDesignadas = secoes.find(s => s.cfgOriginal === CFG_AUDIENCIAS_DESIGNADAS);
        const secaoAudienciasRealizadas = secoes.find(s => s.cfgOriginal === CFG_AUDIENCIAS_REALIZADAS);
        const secaoApreensoes = secoes.find(s => s.cfgOriginal === CFG_APREENSOES);

        // Processos Ativos entra como a primeira linha do Cartório (pedido do usuário) —
        // não é mais um card à parte no topo da capa.
        const mapaAtivos = lerMapaAtivos();
        const atuacoesAtivas = Object.keys(mapaAtivos);
        const linhasCartorio = [];
        if (atuacoesAtivas.length) {
            const totalAtivos = atuacoesAtivas.reduce((s, k) => s + (mapaAtivos[k] || 0), 0);
            linhasCartorio.push({
                nome: 'Processos Ativos',
                indicador: `${totalAtivos} ativo(s)`,
                detalhamento: atuacoesAtivas.length > 1
                    ? atuacoesAtivas.map(a => `${a}: ${mapaAtivos[a]}`).join(' · ')
                    : atuacoesAtivas[0],
                situacaoLabel: '', corTexto: '', semSituacao: true, cfgOriginal: null,
            });
        }
        linhasCartorio.push(...itensCartorio.map(t => ({
            nome: t.secao.cfgOriginal === CFG_RETORNO ? 'Retorno de Conclusão' : t.rotulo,
            indicador: `${t.pendentes} pendente(s)${t.prioritarios ? ` · ${t.prioritarios} prior.` : ''}`,
            detalhamento: t.maisAntiga != null ? `Mais antiga: ${t.maisAntiga} dia(s)` : '—',
            situacaoLabel: (SITUACAO_INFO[t.status] || SITUACAO_INFO.regular).rotulo,
            corTexto: (SITUACAO_INFO[t.status] || SITUACAO_INFO.regular).cor,
            semSituacao: false,
            cfgOriginal: t.secao.cfgOriginal,
        })));

        if (secaoTempoMedio) {
            const validos = secaoTempoMedio.dados.filter(d => d.dias != null);
            const media = validos.length ? validos.reduce((s, d) => s + d.dias, 0) / validos.length : null;
            let periodoTxt = '';
            {
                const per = desembrulharObjeto(store.getItem('projudi_tempomedio_periodo'));
                if (per && (per.ini || per.fim)) periodoTxt = `${per.ini || '?'} a ${per.fim || '?'}`;
            }
            linhasCartorio.push({
                nome: 'Tempo Médio de Cumprimento',
                indicador: media != null ? `${media.toFixed(1).replace('.', ',')} dia(s) méd.` : '—',
                detalhamento: periodoTxt ? `Período: ${periodoTxt}` : '—',
                situacaoLabel: '', corTexto: '', semSituacao: true, cfgOriginal: CFG_TEMPOMEDIO,
            });
        }
        if (secaoAudiencias) {
            let detalhamento = '—';
            let mais = null;
            secaoAudiencias.dados.forEach(d => {
                const ts = parseDataBR(d.dataAudiencia);
                if (ts != null && (!mais || ts < mais.ts)) mais = { ts, data: d.dataAudiencia };
            });
            if (mais) {
                const dias = Math.round((now - mais.ts) / DIA_MS);
                detalhamento = `Mais antiga sem termo: ${mais.data} (${dias} dia(s))`;
            }
            linhasCartorio.push({
                nome: 'Audiências Pendentes',
                indicador: `${secaoAudiencias.dados.length} audiência(s)`,
                detalhamento, situacaoLabel: '', corTexto: '', semSituacao: true, cfgOriginal: CFG_AUDIENCIAS,
            });
        }
        if (secaoAudienciasDesignadas) {
            const r = secaoAudienciasDesignadas.dados[0] || {};
            const tsUltima = r.ultimaData ? parseDataBR(r.ultimaData) : null;
            const diasAteUltima = tsUltima != null ? Math.round((tsUltima - now) / DIA_MS) : null;
            const statusAD = classificarSituacaoPorDias(diasAteUltima, 180, 360);
            const infoAD = SITUACAO_INFO[statusAD] || SITUACAO_INFO.regular;
            linhasCartorio.push({
                nome: 'Audiências Designadas',
                indicador: `${r.totalDesignadas || 0} designada(s)`,
                detalhamento: r.ultimaData ? `Última: ${r.ultimaData} (${diasAteUltima} dia(s))` : '—',
                situacaoLabel: infoAD.rotulo, corTexto: infoAD.cor, semSituacao: false, cfgOriginal: CFG_AUDIENCIAS_DESIGNADAS,
            });
        }
        if (secaoAudienciasRealizadas) {
            const r = secaoAudienciasRealizadas.dados[0] || {};
            linhasCartorio.push({
                nome: 'Audiências Realizadas',
                indicador: `${r.totalGeral || 0} realizada(s)`,
                detalhamento: `${r.canceladas || 0} cancel. · ${r.naoRealizadas || 0} não real. · ${r.redesignadas || 0} redesig.`,
                situacaoLabel: '', corTexto: '', semSituacao: true, cfgOriginal: CFG_AUDIENCIAS_REALIZADAS,
            });
        }
        // "Bens Apreendidos" — linha do Cartório (pedido do usuário), mesmo sendo um
        // relatório da categoria Crime. O indicador é só a contagem total (pedido:
        // "apenas o número de apreensões" no quadro); o detalhamento compacta a
        // distribuição por Tipo da apreensão (top 5 + "Outros", mesmo corte usado nos
        // gráficos — ver contarPorCampo), sem estourar a célula.
        if (secaoApreensoes) {
            const porTipo = contarPorCampo(secaoApreensoes.dados, 'tipo', 5);
            const detalhamento = porTipo.map(it => `${it.label}: ${it.valor}`).join(' · ') || '—';
            linhasCartorio.push({
                nome: 'Bens Apreendidos',
                indicador: `${secaoApreensoes.dados.length} apreensão(ões)`,
                detalhamento,
                situacaoLabel: '', corTexto: '', semSituacao: true, cfgOriginal: CFG_APREENSOES,
            });
        }
        // Extração pulada pelo usuário (ver pularRelatorioAtual): sobrepõe o que quer que
        // tenha sido calculado acima — o dado pode estar incompleto, então avisa em vez de
        // fingir que terminou normalmente.
        linhasCartorio.forEach(l => {
            if (l.cfgOriginal && foiInterrompidoPorErro(l.cfgOriginal)) {
                l.detalhamento = 'Extração interrompida por erro — dados parciais';
                l.situacaoLabel = 'Interrompido'; l.corTexto = COR.vermelho; l.semSituacao = false;
            }
        });
        cartorio.linhas = linhasCartorio;

        let gabinete = { itens: [], totalPendentes: 0, totalPrioritarios: 0, situacao: 'regular' };
        if (secaoGabinete) {
            const porJuiz = new Map();
            secaoGabinete.dados.forEach(d => {
                const nome = (d.responsavel || '').trim() || '(sem responsável)';
                if (!porJuiz.has(nome)) porJuiz.set(nome, []);
                porJuiz.get(nome).push(d);
            });
            const itensGabinete = [...porJuiz.entries()].map(([nome, sub]) => {
                const itens = sub.map(d => ({ dias: diasNum(d.dtRemessa, now), prioritario: !!d.prioritario }));
                const maisAntiga = maiorDias(itens);
                return {
                    rotulo: nome, dados: sub,
                    pendentes: sub.length,
                    prioritarios: contarPrioritarios(sub),
                    maisAntiga,
                    status: classificarSituacaoPorDias(maisAntiga, LIMITES_GABINETE.atencao, LIMITES_GABINETE.critico),
                };
            }).sort((a, b) => b.pendentes - a.pendentes);
            gabinete = {
                itens: itensGabinete,
                totalPendentes: secaoGabinete.dados.length,
                totalPrioritarios: contarPrioritarios(secaoGabinete.dados),
                situacao: piorSituacao(itensGabinete.map(it => it.status)),
            };
        }

        const temConteudo = itensCartorio.length > 0 || gabinete.itens.length > 0 || atuacoesAtivas.length > 0;
        let usouPagina1 = false;

        // ═══ SITUAÇÃO DA UNIDADE — processos ativos por atuação, depois Cartório e
        // Gabinete. Se o conteúdo passar de uma página, segue normalmente na próxima. ═══
        if (temConteudo) {
            const primeira = !usouPagina1;
            const pgCapa = doc.internal.getNumberOfPages() + (primeira ? 0 : 1);
            desenharCapaSituacao(doc, cartorio, gabinete, mapaAtivos, agora, primeira);
            doc.outline.add(null, 'Situação da Unidade', { pageNumber: pgCapa });
            usouPagina1 = true;
        }

        const pw = doc.internal.pageSize.getWidth();
        const ph = doc.internal.pageSize.getHeight();

        // ═══ PASSO 1: todos os RESUMOS (Cartório, depois Gabinete, depois outras) ═══
        // As tabelas discriminadas só vêm depois de TODOS os resumos — por isso o link
        // "Ver tabela detalhada →" de cada resumo só pode ser desenhado no passo 3,
        // quando a página real da tabela já é conhecida.
        // Zerado (0 pendências) não gera página de resumo própria — o número já consta na
        // tabela unificada da capa, e um resumo detalhado (gráficos vazios etc.) não
        // acrescenta nada (pedido do usuário).
        let bmCartorio = null;
        itensCartorio.filter(t => t.pendentes > 0).forEach(t => {
            const primeira = !usouPagina1;
            const pg = doc.internal.getNumberOfPages() + (primeira ? 0 : 1);
            usouPagina1 = true;
            t.secao.montarResumo(doc, t.dados, primeira, false);
            t.pgResumoInicio = pg;
            t.pgResumoFim = doc.internal.getNumberOfPages();
            if (!bmCartorio) bmCartorio = doc.outline.add(null, 'Cartório', { pageNumber: pg });
            doc.outline.add(bmCartorio, `${t.rotulo} (${t.pendentes})`, { pageNumber: pg });
        });

        let bmGabinete = null;
        gabinete.itens.forEach(info => {
            const primeira = !usouPagina1;
            const pg = doc.internal.getNumberOfPages() + (primeira ? 0 : 1);
            usouPagina1 = true;
            montarResumoJuizConclusoes(doc, info.rotulo, info.dados, now, primeira);
            info.pgResumoInicio = pg;
            info.pgResumoFim = doc.internal.getNumberOfPages();
            if (!bmGabinete) bmGabinete = doc.outline.add(null, 'Gabinete', { pageNumber: pg });
            info._bmJuiz = doc.outline.add(bmGabinete, `${info.rotulo} (${info.pendentes})`, { pageNumber: pg });
        });

        // Mesma dispensa do Cartório para os relatórios "fora do esquema": zerado só
        // aparece na tabela unificada, sem resumo próprio.
        function secaoVazia(s) {
            if (s.cfgOriginal === CFG_AUDIENCIAS_DESIGNADAS) return !(s.dados[0] && s.dados[0].totalDesignadas > 0);
            if (s.cfgOriginal === CFG_AUDIENCIAS_REALIZADAS) return !(s.dados[0] && s.dados[0].totalGeral > 0);
            return s.dados.length === 0;
        }
        outrasSecoes.filter(s => !secaoVazia(s)).forEach(s => {
            const primeira = !usouPagina1;
            const pg = doc.internal.getNumberOfPages() + (primeira ? 0 : 1);
            usouPagina1 = true;
            s.montarResumo(doc, s.dados, primeira, false);
            s.pgResumoInicio = pg;
            s.pgResumoFim = doc.internal.getNumberOfPages();
            s._bmOutra = doc.outline.add(null, s.rotulo, { pageNumber: pg });
        });

        // ═══ PASSO 2: todas as TABELAS (mesma ordem), cada uma já com o link "← Voltar
        // ao resumo" (a página do resumo já é conhecida, desenhada no passo 1) ═══
        if (!somenteResumo) {
            // Sem pendências, não há o que discriminar — dispensa a tabela (mas o resumo com
            // o KPI já foi desenhado no passo 1 de qualquer forma, ex.: Suspensos zerado).
            itensCartorio.filter(t => t.pendentes > 0).forEach(t => {
                t.pgTabela = t.secao.montarTabela(doc, t.dados, { label: '← Voltar ao resumo', pageNumber: t.pgResumoInicio });
                doc.outline.add(bmCartorio, `${t.rotulo} — tabela detalhada`, { pageNumber: t.pgTabela });
            });
            gabinete.itens.forEach(info => {
                info.pgTabela = montarTabelaJuizConclusoes(doc, info.rotulo, info.dados, now,
                    { label: '← Voltar ao resumo', pageNumber: info.pgResumoInicio });
                doc.outline.add(info._bmJuiz, 'Tabela detalhada', { pageNumber: info.pgTabela });
            });
            // Mesma dispensa do Cartório: sem registros, não há tabela discriminada a
            // mostrar (ex.: Audiências Pendentes zerado).
            outrasSecoes.filter(secaoTemTabela).forEach(s => {
                s.pgTabela = s.montarTabela(doc, s.dados, { label: '← Voltar ao resumo', pageNumber: s.pgResumoInicio });
                doc.outline.add(s._bmOutra, 'Tabela detalhada', { pageNumber: s.pgTabela });
            });

            // ═══ PASSO 3: volta a cada resumo e desenha "Ver tabela detalhada →" agora
            // que a página da tabela é conhecida ═══
            itensCartorio.filter(t => t.pendentes > 0).forEach(t => {
                doc.setPage(t.pgResumoFim);
                desenharLinkRodape(doc, 'Ver tabela detalhada →', t.pgTabela, pw, ph);
            });
            gabinete.itens.forEach(info => {
                doc.setPage(info.pgResumoFim);
                desenharLinkRodape(doc, 'Ver tabela detalhada →', info.pgTabela, pw, ph);
            });
            outrasSecoes.filter(secaoTemTabela).forEach(s => {
                doc.setPage(s.pgResumoFim);
                desenharLinkRodape(doc, 'Ver tabela detalhada →', s.pgTabela, pw, ph);
            });
        }

        // ═══ PASSO 4: transforma cada linha da tabela unificada do Cartório (na capa) num
        // link clicável até a página de resumo daquele relatório — agora que todas as
        // páginas de resumo já são conhecidas (ver _rect capturado em
        // desenharBlocoCartorioUnificado). Linhas sem seção própria (ex.: "Processos
        // Ativos") ou cuja seção foi zerada/pulada e não gerou página ficam sem link.
        cartorio.linhas.forEach(l => {
            if (!l._rect || !l.cfgOriginal) return;
            const t = itensCartorio.find(it => it.secao.cfgOriginal === l.cfgOriginal);
            const s = outrasSecoes.find(o => o.cfgOriginal === l.cfgOriginal);
            const pgAlvo = (t && t.pgResumoInicio) || (s && s.pgResumoInicio);
            if (!pgAlvo) return;
            doc.setPage(l._rect.page);
            doc.link(l._rect.x, l._rect.y, l._rect.w, l._rect.h, { pageNumber: pgAlvo });
        });

        const sufixo = somenteResumo ? '_resumo' : '';
        baixarBlob(doc.output('blob'), `relatorio_conjunto_projudi${sufixo}_${dataArquivo()}.pdf`);
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
        const areaH = Math.max(6, h - 10);
        if (!itens.length) return;
        const labelW = Math.min(w * 0.62, 120);
        const barX = x + labelW;
        const barMaxW = Math.max(6, w - labelW - 14);
        const maxVal = Math.max(...itens.map(i => i.dias)) || 1;
        const linhaH = areaH / itens.length;
        const barH = Math.max(1.6, linhaH * 0.3);
        // Fonte e offsets das duas linhas de rótulo (processo + classe) acompanham a altura
        // da linha, para nunca invadir a linha seguinte quando o gráfico fica espremido.
        const fonteProcesso = Math.max(5.3, Math.min(7.5, linhaH * 0.46));
        const fonteClasse = Math.max(4.8, Math.min(6.8, linhaH * 0.4));
        const temDuasLinhas = linhaH >= (fonteProcesso + fonteClasse) * 1.15;

        itens.forEach((it, i) => {
            const rowY = topo + i * linhaH;
            const cor = it.prioritario ? COR_PRIORITARIO : COR.tinta;
            if (temDuasLinhas) {
                const yProc = rowY + fonteProcesso * 0.62;
                doc.setFont('PublicSans', 'bold'); doc.setFontSize(fonteProcesso); doc.setTextColor(...cor);
                doc.text(doc.splitTextToSize(it.processo || '', labelW - 3)[0], x, yProc);
                doc.setFont('PublicSans', 'normal'); doc.setFontSize(fonteClasse); doc.setTextColor(...COR.muted);
                doc.text(doc.splitTextToSize(it.classe || '', labelW - 3)[0], x, yProc + fonteClasse * 0.95 + 1);
            } else {
                // Sem espaço para duas linhas: mostra só o processo, centralizado na linha.
                doc.setFont('PublicSans', 'bold'); doc.setFontSize(fonteProcesso); doc.setTextColor(...cor);
                doc.text(doc.splitTextToSize(it.processo || '', labelW - 3)[0], x, rowY + linhaH / 2 + fonteProcesso * 0.32);
            }

            const meio = rowY + linhaH / 2;
            const bw = Math.max(0.6, (it.dias / maxVal) * barMaxW);
            doc.setFillColor(...(it.prioritario ? COR_PRIORITARIO : COR.azul));
            doc.roundedRect(barX, meio - barH / 2, bw, barH, 0.5, 0.5, 'F');
            doc.setFont('PublicSans', 'bold'); doc.setFontSize(fonteProcesso); doc.setTextColor(...COR.tinta);
            doc.text(`${it.dias} dia${it.dias === 1 ? '' : 's'}`, barX + bw + 2, meio + fonteProcesso * 0.16);
        });
    }

    const TITULO_TEMPOMEDIO = 'Tempo Médio de Cumprimento';

    function gerarPDFTempoMedio(dados, somenteResumo) {
        const doc = novoDocPDF();
        montarResumoTempoMedio(doc, dados, true, false);
        doc.outline.add(null, 'Resumo', { pageNumber: 1 });
        if (!somenteResumo) {
            const pgTabela = montarTabelaTempoMedio(doc, dados, false);
            doc.outline.add(null, 'Tabela detalhada', { pageNumber: pgTabela });
        }
        const sufixo = somenteResumo ? '_resumo' : '';
        baixarBlob(doc.output('blob'), `tempo_medio_projudi${sufixo}_${dataArquivo()}.pdf`);
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
        let fimPeriodo = '';
        {
            // Mesma proteção contra JSON codificado em camadas usada em lerFilaMesesTempoMedio.
            let v = store.getItem('projudi_tempomedio_periodo') || '{}';
            let t = 0;
            while (typeof v === 'string' && t < 5) { try { v = JSON.parse(v); } catch (e) { v = {}; break; } t++; }
            const per = (v && typeof v === 'object' && !Array.isArray(v)) ? v : {};
            if (per.ini || per.fim) periodoStr = `${per.ini || '?'} a ${per.fim || '?'}`;
            fimPeriodo = per.fim || '';
        }
        // "Até dd/mm" (sem o ano — o termo final é sempre dentro do período já mostrado no
        // subtítulo) usado no rótulo do KPI de não cumpridas, ver abaixo.
        const fimCurto = fimPeriodo ? fimPeriodo.slice(0, 5) : '';

        // Decisões ainda não cumpridas (dtCartorio vazia = cartório não analisou ainda)
        const naoCumpridas = dados.filter(d => !d.dtCartorio);
        const maisAntigaNC = naoCumpridas.reduce((best, d) => {
            const ts = parseDataBR(d.dtAnalise);
            return ts != null && (best === null || ts < best.ts) ? { ts, str: d.dtAnalise } : best;
        }, null);

        doc.setFont('PublicSans', 'bold'); doc.setFontSize(16); doc.setTextColor(...COR.tinta);
        doc.text(TITULO_TEMPOMEDIO, m, m + 2);
        doc.setFont('PublicSans', 'normal'); doc.setFontSize(9); doc.setTextColor(...COR.tintaSec);
        let subtitulo = `Extraído em ${hoje} às ${hora}  •  ${dados.length} registro(s) analisado(s)`;
        if (periodoStr) subtitulo += `  •  Período: ${periodoStr}`;
        const fraseCompTM = fraseCompetencias(dados);
        if (fraseCompTM) subtitulo += `  •  ${fraseCompTM}`;
        const linhasSubtitulo = doc.splitTextToSize(subtitulo, uw);
        doc.text(linhasSubtitulo, m, m + 8);
        const yLinhaTM = m + 8 + (linhasSubtitulo.length - 1) * 4.2 + 3;
        doc.setDrawColor(...COR.azul); doc.setLineWidth(0.5); doc.line(m, yLinhaTM, pw - m, yLinhaTM);

        const gap = 6;
        // Linha 1: 4 KPIs centralizados — analisados / tempo médio geral / prioritários / não cumpridas
        const kY = yLinhaTM + 5;
        const kW4 = (uw - 3 * gap) / 4;
        const prioPct = validos.length ? Math.round(prioritarios.length / validos.length * 100) : 0;
        desenharCard(doc, m,               kY, kW4, 28, 'Registros analisados', String(dados.length), [], true, COR.azul);
        desenharCard(doc, m + kW4 + gap,   kY, kW4, 28, 'Tempo médio geral', fmtDias(geral), [], true, COR.aqua);
        desenharCard(doc, m + 2*(kW4+gap), kY, kW4, 28, 'Prioritários', String(prioritarios.length), [`${prioPct}% do total`], true, COR.vermelho);
        desenharCard(doc, m + 3*(kW4+gap), kY, kW4, 28, `Não cumpridas${fimCurto ? ` (até ${fimCurto})` : ''}`, String(naoCumpridas.length),
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

        // ═══ PÁGINA 2 (só quando há dados): Tempo médio por Usuário do Cartório e por
        // Tipo de Conclusão — o resumo agora pode passar de uma página quando esses
        // gráficos complementares não cabem na primeira (que já vem cheia com os 4+2+1
        // KPIs e os dois gráficos anteriores). Login extraído da própria célula "Dt.
        // Análise Cartório" (ver usuarioDeTexto) — o Projudi não expõe essa coluna
        // separada; "Tipo de Conclusão" já vem numa coluna própria (ex.: DECISÃO,
        // DESPACHO, SENTENÇA). Quantidade de atos entre parênteses no rótulo em ambos
        // (agregarMedia já calcula "n" — só falta compor o texto).
        const porUsuario = agregarMedia(validos, 'usuarioCartorio', 'dias', 20)
            .map(it => ({ ...it, label: `${it.label} (${it.n})` }));
        const porTipo = agregarMedia(validos, 'tipoConclusao', 'dias', 12)
            .map(it => ({ ...it, label: `${it.label} (${it.n})` }));
        if (porUsuario.length || porTipo.length) {
            doc.addPage();
            let hy2 = m + 2;
            doc.setFont('PublicSans', 'bold'); doc.setFontSize(16); doc.setTextColor(...COR.tinta);
            doc.text(TITULO_TEMPOMEDIO, m, hy2);
            hy2 += 8;
            doc.setFont('PublicSans', 'normal'); doc.setFontSize(9); doc.setTextColor(...COR.tintaSec);
            doc.text('Gráficos complementares', m, hy2);
            hy2 += 3;
            doc.setDrawColor(...COR.azul); doc.setLineWidth(0.5); doc.line(m, hy2, pw - m, hy2);

            const gY0b = hy2 + 6;
            const disponivel2 = ph - m - gY0b - 14;
            if (porUsuario.length && porTipo.length) {
                // Divide o espaço entre os dois, dando mais altura ao que tem mais itens
                // (mesma lógica de proporção usada nos dois gráficos da página 1).
                const chartGap2 = 8;
                const totalItens = porUsuario.length + porTipo.length;
                const alturaA = Math.max(30, (disponivel2 - chartGap2) * (porUsuario.length / totalItens));
                const alturaB = Math.max(30, disponivel2 - chartGap2 - alturaA);
                desenharBarras(doc, m, gY0b, uw, alturaA, 'Tempo médio por Usuário (Cartório)', porUsuario, fmtDias, COR.aqua);
                desenharBarras(doc, m, gY0b + alturaA + chartGap2, uw, alturaB, 'Tempo médio por Tipo de Conclusão', porTipo, fmtDias, COR.azul);
            } else if (porUsuario.length) {
                desenharBarras(doc, m, gY0b, uw, disponivel2, 'Tempo médio por Usuário (Cartório)', porUsuario, fmtDias, COR.aqua);
            } else {
                desenharBarras(doc, m, gY0b, uw, disponivel2, 'Tempo médio por Tipo de Conclusão', porTipo, fmtDias, COR.azul);
            }
            desenharRodape(doc, TITULO_TEMPOMEDIO, `${hoje} ${hora}`, pw, ph, m, comIndice);
        }
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
        tituloSecao(doc, m, m + 3, pw - 2 * m, 'Tabela discriminada — Tempo Médio de Cumprimento');

        const colunas = [
            { header: 'Processo', width: 28, get: (d) => d.processo },
            { header: 'Dt. Análise', width: 20, get: (d) => d.dtAnalise },
            { header: 'Dt. Análise Cartório', width: 20, get: (d) => d.dtCartorio },
            { header: 'Usuário Cartório', width: 20, get: (d) => d.usuarioCartorio },
            { header: 'Tipo de Conclusão', width: 28, get: (d) => d.tipoConclusao },
            { header: 'Classe Processual', width: 40, get: (d) => d.classe },
            { header: 'Dias p/ Cumprimento', width: 14, get: (d) => (d.dias == null ? '' : String(d.dias)) },
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
            styles: { font: 'PublicSans', fontSize: 7.5, cellPadding: 1.6, textColor: COR.tintaSec,
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

    // ── PDF do relatório de Audiências Pendentes (primeiro item específico da
    // categoria Crime) ───────────────────────────────────────────────────────────

    const TITULO_AUDIENCIAS = 'Audiências Pendentes';

    function gerarPDFAudiencias(dados, somenteResumo) {
        const doc = novoDocPDF();
        montarResumoAudiencias(doc, dados, true, false, somenteResumo);
        doc.outline.add(null, 'Resumo', { pageNumber: 1 });
        const sufixo = somenteResumo ? '_resumo' : '';
        baixarBlob(doc.output('blob'), `audiencias_pendentes_projudi${sufixo}_${dataArquivo()}.pdf`);
    }

    // Resumo (KPIs + gráfico) seguido, sem quebra de seção, pela tabela discriminada —
    // sempre juntos, nunca com a tabela numa seção separada atrás de um link "Ver tabela
    // detalhada →" (pedido do usuário). Continua na mesma página quando cabe; senão, o
    // autoTable segue para quantas páginas forem precisas, normalmente. somenteResumo
    // omite a tabela (usado tanto no relatório individual quanto no Relatório PDF conjunto,
    // pela opção "Só resumo").
    function montarResumoAudiencias(doc, dados, ehPrimeiraSecao, comIndice, somenteResumo) {
        if (!ehPrimeiraSecao) doc.addPage();
        const agora = new Date();
        const pw = doc.internal.pageSize.getWidth();
        const ph = doc.internal.pageSize.getHeight();
        const m = 12;
        const uw = pw - 2 * m;
        const hoje = agora.toLocaleDateString('pt-BR');
        const hora = agora.toLocaleTimeString('pt-BR', { hour: '2-digit', minute: '2-digit' });

        const prioritarios = dados.filter(d => d.prioritario);
        const prioPct = dados.length ? Math.round(prioritarios.length / dados.length * 100) : 0;

        doc.setFont('PublicSans', 'bold'); doc.setFontSize(16); doc.setTextColor(...COR.tinta);
        doc.text(TITULO_AUDIENCIAS, m, m + 2);
        doc.setFont('PublicSans', 'normal'); doc.setFontSize(9); doc.setTextColor(...COR.tintaSec);
        let subtitulo = `Extraído em ${hoje} às ${hora}  •  ${dados.length} audiência(s) pendente(s)`;
        const fraseComp = fraseCompetencias(dados);
        if (fraseComp) subtitulo += `  •  ${fraseComp}`;
        const linhasSubtitulo = doc.splitTextToSize(subtitulo, uw);
        doc.text(linhasSubtitulo, m, m + 8);
        const yLinha = m + 8 + (linhasSubtitulo.length - 1) * 4.2 + 3;
        doc.setDrawColor(...COR.azul); doc.setLineWidth(0.5); doc.line(m, yLinha, pw - m, yLinha);

        const gap = 6;
        const kY = yLinha + 5;
        const kW2 = (uw - gap) / 2;
        desenharCard(doc, m,             kY, kW2, 28, 'Audiências Pendentes', String(dados.length), [], true, COR.azul);
        desenharCard(doc, m + kW2 + gap, kY, kW2, 28, 'Prioritárias', String(prioritarios.length), [`${prioPct}% do total`], true, COR.vermelho);

        // Altura FIXA para o gráfico (não consome o resto da página) — deixa espaço para a
        // tabela discriminada logo abaixo, na mesma página quando couber.
        const chartY = kY + 28 + gap + 2;
        const alturaChart = 46;
        const porTipo = contarPorCampo(dados, 'tipoAudiencia', 12);
        let proximoY = chartY;
        if (porTipo.length) {
            desenharBarras(doc, m, chartY, uw, alturaChart, 'Audiências por Tipo', porTipo, undefined, COR.aqua);
            proximoY = chartY + alturaChart + 6;
        }

        if (somenteResumo || !dados.length) {
            desenharRodape(doc, TITULO_AUDIENCIAS, `${hoje} ${hora}`, pw, ph, m, comIndice);
            return;
        }

        // Da audiência mais próxima para a mais distante (sem data válida vai ao final) —
        // mais útil que a ordem de coleta para saber o que vem primeiro.
        const ordenados = dados.slice().sort((a, b) => {
            const ta = parseDataBR(a.dataAudiencia); const tb = parseDataBR(b.dataAudiencia);
            return (ta == null ? Infinity : ta) - (tb == null ? Infinity : tb);
        });

        tituloSecao(doc, m, proximoY, uw, 'Detalhamento — Audiências Pendentes');

        const colunas = [
            { header: 'Processo', width: 44, get: (d) => d.processo },
            { header: 'Tipo da Audiência', width: 62, get: (d) => d.tipoAudiencia },
            { header: 'Data da Audiência', width: 36, get: (d) => d.dataAudiencia },
            { header: 'Dias até a Audiência', width: 24, get: (d) => fmtDiferencaDias(diasAteAudiencia(d.dataAudiencia, agora.getTime())) },
        ];
        const columnStyles = {};
        colunas.forEach((c, i) => { columnStyles['k' + i] = { cellWidth: c.width }; });

        doc.autoTable({
            columns: colunas.map((c, i) => ({ header: c.header, dataKey: 'k' + i })),
            body: ordenados.map(d => {
                const o = {};
                colunas.forEach((c, i) => { o['k' + i] = String(c.get(d) ?? ''); });
                return o;
            }),
            startY: proximoY + 6,
            margin: { left: m, right: m, top: m, bottom: 14 },
            theme: 'grid',
            styles: { font: 'PublicSans', fontSize: 8, cellPadding: 2, textColor: COR.tintaSec,
                      lineColor: COR.grade, lineWidth: 0.1, overflow: 'linebreak', valign: 'middle' },
            headStyles: { fillColor: COR.azul, textColor: [255, 255, 255], fontStyle: 'bold', fontSize: 8.5 },
            alternateRowStyles: { fillColor: COR.cartao },
            columnStyles,
            didParseCell: (data) => {
                if (data.section === 'body' && data.column.index === 0 && ordenados[data.row.index] && ordenados[data.row.index].prioritario) {
                    data.cell.styles.textColor = COR_PRIORITARIO;
                    data.cell.styles.fontStyle = 'bold';
                }
            },
            didDrawPage: () => desenharRodape(doc, TITULO_AUDIENCIAS, `${hoje} ${hora}`, pw, ph, m, comIndice),
        });
    }

    // ── PDF do relatório de Audiências Designadas (Crime — Pauta de Horários) ───
    // Os dados coletados já vêm como um resumo pronto (não uma lista de linhas — ver
    // coletarAudienciasDesignadas), com a tabela discriminada guardada em resumo.tabela.

    const TITULO_AUDIENCIAS_DESIGNADAS = 'Audiências Designadas';

    function gerarPDFAudienciasDesignadas(dados) {
        const doc = novoDocPDF();
        const resumo = dados && dados[0] ? dados[0] : null;
        montarResumoAudienciasDesignadas(doc, resumo, true, false);
        doc.outline.add(null, 'Resumo', { pageNumber: 1 });
        if (resumo && resumo.tabela && resumo.tabela.length) {
            const pgTabela = doc.internal.getNumberOfPages() + 1;
            montarTabelaAudienciasDesignadas(doc, resumo.tabela, false);
            doc.outline.add(null, 'Tabela detalhada', { pageNumber: pgTabela });
        }
        baixarBlob(doc.output('blob'), `audiencias_designadas_projudi_${dataArquivo()}.pdf`);
    }

    // Lista de processos formatada para caber numa sub-linha do card (trunca com "+N" se
    // não couber tudo) — usada tanto no card combinado do último dia quanto no card por tipo.
    function fraseProcessos(processos, maxItens) {
        if (!processos || !processos.length) return '';
        const visiveis = processos.slice(0, maxItens);
        const resto = processos.length - visiveis.length;
        return visiveis.join(', ') + (resto > 0 ? ` (+${resto})` : '');
    }

    function montarResumoAudienciasDesignadas(doc, resumo, ehPrimeiraSecao, comIndice) {
        if (!ehPrimeiraSecao) doc.addPage();
        const agora = new Date();
        const pw = doc.internal.pageSize.getWidth();
        const ph = doc.internal.pageSize.getHeight();
        const m = 12;
        const uw = pw - 2 * m;
        const hoje = agora.toLocaleDateString('pt-BR');
        const hora = agora.toLocaleTimeString('pt-BR', { hour: '2-digit', minute: '2-digit' });

        const r = resumo || { totalDesignadas: 0, ultimaData: null, processosUltimoDia: [], totalProcessosUltimoDia: 0, porTipo: [], tabela: [], concentracaoDiaSemana: [], competencia: '' };

        doc.setFont('PublicSans', 'bold'); doc.setFontSize(16); doc.setTextColor(...COR.tinta);
        doc.text(TITULO_AUDIENCIAS_DESIGNADAS, m, m + 2);
        doc.setFont('PublicSans', 'normal'); doc.setFontSize(9); doc.setTextColor(...COR.tintaSec);
        let subtitulo = `Extraído em ${hoje} às ${hora}  •  Pauta de Horários (todos os tipos, 10 anos à frente)`;
        if (r.competencia) subtitulo += `  •  Competência: ${r.competencia}`;
        const linhasSubtitulo = doc.splitTextToSize(subtitulo, uw);
        doc.text(linhasSubtitulo, m, m + 8);
        let yObs = m + 8 + (linhasSubtitulo.length - 1) * 4.2 + 4.2;
        // Observação sobre a junção de tipos (pedido do usuário) — sempre presente, já que
        // a normalização (ver normalizarTipoAudiencia) vale para toda a pauta.
        doc.setFont('PublicSans', 'italic'); doc.setFontSize(7.8); doc.setTextColor(...COR.muted);
        const linhasObs = doc.splitTextToSize(OBSERVACAO_TIPOS_AD, uw);
        doc.text(linhasObs, m, yObs);
        const yLinha = yObs + (linhasObs.length - 1) * 3.4 + 3.5;
        doc.setDrawColor(...COR.azul); doc.setLineWidth(0.5); doc.line(m, yLinha, pw - m, yLinha);

        const gap = 6;
        const kY = yLinha + 5;
        const kW2 = (uw - gap) / 2;
        desenharCard(doc, m,             kY, kW2, 28, 'Total de Audiências Designadas', String(r.totalDesignadas), [], true, COR.azul);
        desenharCard(doc, m + kW2 + gap, kY, kW2, 28, 'Último dia com audiência', r.ultimaData || '—', [], true, COR.aqua);

        // Dois KPIs (pedido do usuário): 1) só a situação da vara (verde até 180 dias até a
        // audiência mais distante, amarelo de 180 a 360, vermelho acima de 360 — mesma
        // régua Regular/Atenção/Crítico usada nos demais relatórios, só que aqui o "atraso"
        // é a agenda já estar preenchida muito longe no futuro); 2) só os processos do dia
        // mais distante, com os dias até essa data.
        const k2Y = kY + 28 + gap;
        const tsUltima = r.ultimaData ? parseDataBR(r.ultimaData) : null;
        const diasAteUltima = tsUltima != null ? Math.round((tsUltima - agora.getTime()) / DIA_MS) : null;
        const status = classificarSituacaoPorDias(diasAteUltima, 180, 360);
        const infoStatus = SITUACAO_INFO[status] || SITUACAO_INFO.regular;
        const totalUltimoDia = r.totalProcessosUltimoDia != null ? r.totalProcessosUltimoDia : (r.processosUltimoDia || []).length;
        // Lista COMPLETA de processos, sem cortar (pedido do usuário) — o card quebra em
        // várias linhas em vez de truncar, então a altura é calculada antes de desenhar.
        const processosTexto = (r.processosUltimoDia || []).join(', ') || '—';
        const subLinhaProcessos = r.ultimaData && diasAteUltima != null ? `${diasAteUltima} dia(s) até ${r.ultimaData}` : 'Nenhuma audiência designada';
        const alturaCard2 = Math.max(28, medirAlturaCardLista(doc, kW2, processosTexto, true));
        desenharCard(doc, m, k2Y, kW2, alturaCard2, 'Situação da Vara', infoStatus.rotulo,
            [diasAteUltima != null ? `${diasAteUltima} dia(s) até a audiência mais distante da pauta` : 'Sem audiências para calcular'],
            true, infoStatus.cor);
        desenharCardLista(doc, m + kW2 + gap, k2Y, kW2, alturaCard2, `Processos no Dia Mais Distante (${totalUltimoDia})`, processosTexto, subLinhaProcessos, COR.ambar);

        const tY = k2Y + alturaCard2 + gap + 4;
        if (r.porTipo.length) {
            tituloSecao(doc, m, tY, uw, 'Última audiência designada por tipo');
            doc.autoTable({
                columns: [
                    { header: 'Tipo da Audiência', dataKey: 'tipo' },
                    { header: 'Data mais distante', dataKey: 'data' },
                    { header: 'Processo(s)', dataKey: 'processos' },
                ],
                body: r.porTipo.map(it => ({ tipo: it.tipo, data: it.data, processos: (it.processos || []).join(', ') || '—' })),
                startY: tY + 6,
                margin: { left: m, right: m, bottom: 14 },
                theme: 'grid',
                styles: { font: 'PublicSans', fontSize: 8.5, cellPadding: 2.2, textColor: COR.tintaSec, lineColor: COR.grade, lineWidth: 0.1, valign: 'middle' },
                headStyles: { fillColor: COR.azul, textColor: [255, 255, 255], fontStyle: 'bold', fontSize: 8 },
                alternateRowStyles: { fillColor: COR.cartao },
                columnStyles: { tipo: { cellWidth: uw * 0.28 }, data: { cellWidth: uw * 0.17, halign: 'right' }, processos: { cellWidth: uw * 0.55 } },
                didDrawPage: () => desenharRodape(doc, TITULO_AUDIENCIAS_DESIGNADAS, `${hoje} ${hora}`, pw, ph, m, comIndice),
            });
        } else {
            desenharRodape(doc, TITULO_AUDIENCIAS_DESIGNADAS, `${hoje} ${hora}`, pw, ph, m, comIndice);
        }

        // Concentração por dia da semana (só dias úteis — ver concentracaoPorDiaSemana):
        // gráfico com o total de processos por dia + detalhamento por tipo, numa página
        // própria dentro do resumo (a mesma seção do relatório, só que numa página a mais).
        const concentracao = r.concentracaoDiaSemana || [];
        if (concentracao.some(d => d.total > 0)) {
            doc.addPage();
            tituloSecao(doc, m, m + 4, uw, 'Concentração de Audiências por Dia da Semana (dias úteis)');
            const itensBarras = concentracao.map(d => ({ label: d.diaSemana, valor: d.total }));
            desenharBarras(doc, m, m + 8, uw, 44, 'Total de processos por dia da semana', itensBarras, undefined, COR.aqua);

            const linhasDetalhe = [];
            concentracao.forEach(d => d.porTipo.forEach(it => linhasDetalhe.push({ dia: d.diaSemana, tipo: it.tipo, quantidade: it.quantidade })));
            const yTabela = m + 8 + 44 + 6;
            if (linhasDetalhe.length) {
                tituloSecao(doc, m, yTabela, uw, 'Detalhamento por dia da semana e tipo de audiência');
                doc.autoTable({
                    columns: [
                        { header: 'Dia da Semana', dataKey: 'dia' },
                        { header: 'Tipo da Audiência', dataKey: 'tipo' },
                        { header: 'Quantidade', dataKey: 'quantidade' },
                    ],
                    body: linhasDetalhe,
                    startY: yTabela + 6,
                    margin: { left: m, right: m, bottom: 14 },
                    theme: 'grid',
                    styles: { font: 'PublicSans', fontSize: 8.5, cellPadding: 2.2, textColor: COR.tintaSec, lineColor: COR.grade, lineWidth: 0.1, valign: 'middle' },
                    headStyles: { fillColor: COR.aqua, textColor: [255, 255, 255], fontStyle: 'bold', fontSize: 8 },
                    alternateRowStyles: { fillColor: COR.cartao },
                    columnStyles: { dia: { cellWidth: uw * 0.25 }, tipo: { cellWidth: uw * 0.5 }, quantidade: { cellWidth: uw * 0.25, halign: 'right' } },
                    didDrawPage: () => desenharRodape(doc, TITULO_AUDIENCIAS_DESIGNADAS, `${hoje} ${hora}`, pw, ph, m, comIndice),
                });
            } else {
                desenharRodape(doc, TITULO_AUDIENCIAS_DESIGNADAS, `${hoje} ${hora}`, pw, ph, m, comIndice);
            }
        }
    }

    // Tabela discriminada com TODAS as audiências designadas (data/hora/processo/tipo) —
    // uma linha por processo (ver coletarAudienciasDesignadas). Segue o mesmo padrão das
    // demais "montarTabelaX": página nova, cabeçalho + rodapé, devolve o nº da página.
    function montarTabelaAudienciasDesignadas(doc, tabela, comIndice) {
        doc.addPage();
        const agora = new Date();
        const pw = doc.internal.pageSize.getWidth();
        const ph = doc.internal.pageSize.getHeight();
        const m = 12;
        const uw = pw - 2 * m;
        const hoje = agora.toLocaleDateString('pt-BR');
        const hora = agora.toLocaleTimeString('pt-BR', { hour: '2-digit', minute: '2-digit' });
        const pagina = doc.internal.getNumberOfPages();

        tituloSecao(doc, m, m + 4, uw, `Tabela discriminada — ${TITULO_AUDIENCIAS_DESIGNADAS} (${tabela.length} registro(s))`);
        doc.autoTable({
            columns: [
                { header: 'Data', dataKey: 'data' },
                { header: 'Hora', dataKey: 'horario' },
                { header: 'Processo', dataKey: 'processo' },
                { header: 'Tipo da Audiência', dataKey: 'tipoAudiencia' },
            ],
            body: tabela.map(t => ({ data: t.data, horario: t.horario, processo: t.processo || '—', tipoAudiencia: t.tipoAudiencia })),
            startY: m + 10,
            margin: { left: m, right: m, bottom: 14 },
            theme: 'grid',
            styles: { font: 'PublicSans', fontSize: 8, cellPadding: 2, textColor: COR.tintaSec, lineColor: COR.grade, lineWidth: 0.1, valign: 'middle' },
            headStyles: { fillColor: COR.azul, textColor: [255, 255, 255], fontStyle: 'bold', fontSize: 8 },
            alternateRowStyles: { fillColor: COR.cartao },
            columnStyles: { data: { cellWidth: uw * 0.14 }, horario: { cellWidth: uw * 0.1 }, processo: { cellWidth: uw * 0.36 }, tipoAudiencia: { cellWidth: uw * 0.4 } },
            didDrawPage: () => desenharRodape(doc, TITULO_AUDIENCIAS_DESIGNADAS, `${hoje} ${hora}`, pw, ph, m, comIndice),
        });
        return pagina;
    }

    // ── PDF do relatório de Audiências Realizadas (Crime — Estatísticas de
    // Audiência) ─────────────────────────────────────────────────────────────────
    // Os dados coletados já vêm como um resumo pronto (total geral + detalhamento por
    // usuário, já filtrado pelo mínimo — ver finalizarAudienciasRealizadas), sem lista de
    // processos: é um relatório de totais, não de discriminação.

    const TITULO_AUDIENCIAS_REALIZADAS = 'Audiências Realizadas';

    function gerarPDFAudienciasRealizadas(dados) {
        const doc = novoDocPDF();
        const resumo = dados && dados[0] ? dados[0] : null;
        montarResumoAudienciasRealizadas(doc, resumo, true, false);
        doc.outline.add(null, 'Resumo', { pageNumber: 1 });
        baixarBlob(doc.output('blob'), `audiencias_realizadas_projudi_${dataArquivo()}.pdf`);
    }

    function montarResumoAudienciasRealizadas(doc, resumo, ehPrimeiraSecao, comIndice) {
        if (!ehPrimeiraSecao) doc.addPage();
        const agora = new Date();
        const pw = doc.internal.pageSize.getWidth();
        const ph = doc.internal.pageSize.getHeight();
        const m = 12;
        const uw = pw - 2 * m;
        const hoje = agora.toLocaleDateString('pt-BR');
        const hora = agora.toLocaleTimeString('pt-BR', { hour: '2-digit', minute: '2-digit' });

        const r = resumo || { totalGeral: 0, canceladas: 0, naoRealizadas: 0, redesignadas: 0, periodo: { dataInicio: '', dataFim: '' }, porUsuario: [], usuariosIgnorados: 0, totalUsuarios: 0, competencia: '' };
        const periodoTxt = (r.periodo && r.periodo.dataInicio && r.periodo.dataFim) ? `${r.periodo.dataInicio} a ${r.periodo.dataFim}` : '—';

        doc.setFont('PublicSans', 'bold'); doc.setFontSize(16); doc.setTextColor(...COR.tinta);
        doc.text(TITULO_AUDIENCIAS_REALIZADAS, m, m + 2);
        doc.setFont('PublicSans', 'normal'); doc.setFontSize(9); doc.setTextColor(...COR.tintaSec);
        let subtitulo = `Extraído em ${hoje} às ${hora}  •  Período: ${periodoTxt}`;
        if (r.competencia) subtitulo += `  •  Competência: ${r.competencia}`;
        const linhasSubtitulo = doc.splitTextToSize(subtitulo, uw);
        doc.text(linhasSubtitulo, m, m + 8);
        let yObs = m + 8 + (linhasSubtitulo.length - 1) * 4.2 + 4.2;
        doc.setFont('PublicSans', 'italic'); doc.setFontSize(7.8); doc.setTextColor(...COR.muted);
        const obs = `Observação: o detalhamento por usuário mostra apenas quem realizou ${MIN_AUDIENCIAS_POR_USUARIO_AR} ou mais audiências no período `
            + `(${r.usuariosIgnorados} usuário(s) abaixo desse mínimo não aparecem na lista, mas contam no total geral).`;
        const linhasObs = doc.splitTextToSize(obs, uw);
        doc.text(linhasObs, m, yObs);
        const yLinha = yObs + (linhasObs.length - 1) * 3.4 + 3.5;
        doc.setDrawColor(...COR.azul); doc.setLineWidth(0.5); doc.line(m, yLinha, pw - m, yLinha);

        const gap = 6;
        const kY = yLinha + 5;
        const kW2 = (uw - gap) / 2;
        desenharCard(doc, m,             kY, kW2, 28, 'Total de Audiências Realizadas', String(r.totalGeral), [], true, COR.azul);
        desenharCard(doc, m + kW2 + gap, kY, kW2, 28, `Usuários com ${MIN_AUDIENCIAS_POR_USUARIO_AR}+ audiências`, String(r.porUsuario.length), [], true, COR.aqua);

        // Canceladas/Não Realizadas/Redesignadas — extraídas só da pesquisa geral (pedido
        // do usuário), lado a lado numa segunda linha de KPIs menores.
        const k2Y = kY + 28 + gap;
        const kW3 = (uw - 2 * gap) / 3;
        desenharCard(doc, m,                  k2Y, kW3, 24, 'Canceladas', String(r.canceladas), [], true, COR.ambar);
        desenharCard(doc, m + kW3 + gap,      k2Y, kW3, 24, 'Não Realizadas', String(r.naoRealizadas), [], true, COR.vermelho);
        desenharCard(doc, m + 2 * (kW3 + gap), k2Y, kW3, 24, 'Redesignadas', String(r.redesignadas), [], true, COR.muted);

        // Sem gráfico aqui — a mesma informação já está na tabela abaixo (pedido do
        // usuário), com o percentual de cada categoria sobre o total do magistrado(a)
        // (realizadas + canceladas + não realizadas + redesignadas).
        const tY = k2Y + 24 + gap + 4;
        if (r.porUsuario.length) {
            tituloSecao(doc, m, tY, uw, 'Detalhamento por usuário');
            const fmtPct = (n, total) => total > 0 ? `${Math.round(n / total * 100)}%` : '—';
            doc.autoTable({
                columns: [
                    { header: 'Usuário', dataKey: 'nome' },
                    { header: 'Realizadas', dataKey: 'quantidade' },
                    { header: 'Canceladas', dataKey: 'canceladas' },
                    { header: 'Não Realizadas', dataKey: 'naoRealizadas' },
                    { header: 'Redesignadas', dataKey: 'redesignadas' },
                ],
                body: r.porUsuario.map(u => {
                    const canceladas = u.canceladas || 0, naoRealizadas = u.naoRealizadas || 0, redesignadas = u.redesignadas || 0;
                    const totalUsuario = u.quantidade + canceladas + naoRealizadas + redesignadas;
                    return {
                        nome: u.nome,
                        quantidade: `${u.quantidade} (${fmtPct(u.quantidade, totalUsuario)})`,
                        canceladas: `${canceladas} (${fmtPct(canceladas, totalUsuario)})`,
                        naoRealizadas: `${naoRealizadas} (${fmtPct(naoRealizadas, totalUsuario)})`,
                        redesignadas: `${redesignadas} (${fmtPct(redesignadas, totalUsuario)})`,
                    };
                }),
                startY: tY + 6,
                margin: { left: m, right: m, bottom: 14 },
                theme: 'grid',
                styles: { font: 'PublicSans', fontSize: 8.5, cellPadding: 2.2, textColor: COR.tintaSec, lineColor: COR.grade, lineWidth: 0.1, valign: 'middle' },
                headStyles: { fillColor: COR.azul, textColor: [255, 255, 255], fontStyle: 'bold', fontSize: 8 },
                alternateRowStyles: { fillColor: COR.cartao },
                columnStyles: {
                    nome: { cellWidth: uw * 0.32 },
                    quantidade: { cellWidth: uw * 0.17, halign: 'right' },
                    canceladas: { cellWidth: uw * 0.17, halign: 'right' },
                    naoRealizadas: { cellWidth: uw * 0.17, halign: 'right' },
                    redesignadas: { cellWidth: uw * 0.17, halign: 'right' },
                },
                didDrawPage: () => desenharRodape(doc, TITULO_AUDIENCIAS_REALIZADAS, `${hoje} ${hora}`, pw, ph, m, comIndice),
            });
        } else {
            desenharRodape(doc, TITULO_AUDIENCIAS_REALIZADAS, `${hoje} ${hora}`, pw, ph, m, comIndice);
        }
    }

    // ── PDF do relatório de Processos Paralisados ───────────────────────────────
    // Segue o mesmo padrão do Tempo Médio (dias já vêm prontos do Projudi, sem precisar
    // calcular a partir de duas datas): KPIs + top 10 mais tempo paralisados + média por classe.

    const TITULO_PARALISADOS = 'Processos Paralisados';

    function gerarPDFParalisados(dados, somenteResumo) {
        const doc = novoDocPDF();
        montarResumoParalisados(doc, dados, true, false);
        doc.outline.add(null, 'Resumo', { pageNumber: 1 });
        if (!somenteResumo) {
            const pgTabela = montarTabelaParalisados(doc, dados, false);
            doc.outline.add(null, 'Tabela detalhada', { pageNumber: pgTabela });
        }
        const sufixo = somenteResumo ? '_resumo' : '';
        baixarBlob(doc.output('blob'), `paralisados_projudi${sufixo}_${dataArquivo()}.pdf`);
    }

    // Página de RESUMO (KPIs + gráficos) do relatório de Paralisados, dentro de um doc
    // jsPDF já existente. ehPrimeiraSecao=false começa em página nova (uso no conjunto).
    function montarResumoParalisados(doc, dados, ehPrimeiraSecao, comIndice) {
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
        const maisParado = validos.slice().sort((a, b) => b.dias - a.dias)[0] || null;

        doc.setFont('PublicSans', 'bold'); doc.setFontSize(16); doc.setTextColor(...COR.tinta);
        doc.text(TITULO_PARALISADOS, m, m + 2);
        doc.setFont('PublicSans', 'normal'); doc.setFontSize(9); doc.setTextColor(...COR.tintaSec);
        let subtituloParalisados = `Extraído em ${hoje} às ${hora}  •  ${dados.length} processo(s) paralisado(s)`;
        const fraseCompParalisados = fraseCompetencias(dados);
        if (fraseCompParalisados) subtituloParalisados += `  •  ${fraseCompParalisados}`;
        const linhasSubtituloParalisados = doc.splitTextToSize(subtituloParalisados, uw);
        doc.text(linhasSubtituloParalisados, m, m + 8);
        const yLinhaParalisados = m + 8 + (linhasSubtituloParalisados.length - 1) * 4.2 + 3;
        doc.setDrawColor(...COR.azul); doc.setLineWidth(0.5); doc.line(m, yLinhaParalisados, pw - m, yLinhaParalisados);

        const gap = 6;
        // Linha 1: 3 KPIs centralizados — paralisados / tempo médio paralisado / prioritários
        const kY = yLinhaParalisados + 5;
        const kW3 = (uw - 2 * gap) / 3;
        const prioPct = validos.length ? Math.round(prioritarios.length / validos.length * 100) : 0;
        desenharCard(doc, m,               kY, kW3, 28, 'Processos paralisados', String(dados.length), [], true, COR.azul);
        desenharCard(doc, m + kW3 + gap,   kY, kW3, 28, 'Tempo médio paralisado', fmtDias(geral), [], true, COR.aqua);
        desenharCard(doc, m + 2*(kW3+gap), kY, kW3, 28, 'Prioritários', String(prioritarios.length), [`${prioPct}% do total`], true, COR.vermelho);

        // Linha 2: tempo médio paralisado prioritários vs não prioritários (centralizados)
        const k2Y = kY + 28 + gap;
        const kW2 = (uw - gap) / 2;
        desenharCard(doc, m,           k2Y, kW2, 26, 'Tempo médio paralisado — Prioritários',     fmtDias(mediaPrio),    [`${prioritarios.length} processo(s)`], true, COR.vermelho);
        desenharCard(doc, m + kW2+gap, k2Y, kW2, 26, 'Tempo médio paralisado — Não prioritários', fmtDias(mediaNaoPrio), [`${naoPrioritarios.length} processo(s)`], true, COR.azul);

        // Card largo: processo paralisado há mais tempo (centralizado)
        const k3Y = k2Y + 26 + gap;
        let valMP = '—', subsMP = ['Nenhum registro com dados válidos'];
        if (maisParado) {
            valMP = `${maisParado.dias} dia${maisParado.dias === 1 ? '' : 's'}`;
            subsMP = [
                `Processo ${maisParado.processo}${maisParado.prioritario ? '  — PRIORITÁRIO' : ''}`,
                maisParado.classe || '',
            ];
        }
        desenharCard(doc, m, k3Y, uw, 26, 'Processo paralisado há mais tempo', valMP, subsMP, true, COR.vermelho);

        // Gráficos 1 e 2 (largura total, empilhados): dividem o espaço vertical restante da
        // página de forma que a soma das duas alturas + o espaçamento sempre caiba.
        const chartY = k3Y + 26 + gap + 2;
        const chartGap = 8;
        const disponivel = ph - m - chartY - 14; // reserva a faixa do rodapé
        const top10 = validos.slice().sort((a, b) => b.dias - a.dias).slice(0, 10)
            .map(d => ({ processo: d.processo, classe: d.classe, dias: d.dias, prioritario: d.prioritario }));
        const chart1Desejado = Math.max(30, top10.length * 8 + 8);
        const chart1H = Math.min(chart1Desejado, Math.max(30, disponivel - chartGap - 40));
        const chart2H = Math.max(30, disponivel - chart1H - chartGap);
        desenharTopDemorados(doc, m, chartY, uw, chart1H, 'Processos paralisados há mais tempo', top10);

        const chart2Y = chartY + chart1H + chartGap;
        const porClasse = agregarMedia(validos, 'classe', 'dias', 12);
        desenharBarras(doc, m, chart2Y, uw, chart2H, 'Tempo médio paralisado por Classe Processual', porClasse, fmtDias, COR.aqua);

        desenharRodape(doc, TITULO_PARALISADOS, `${hoje} ${hora}`, pw, ph, m, comIndice);
    }

    // Tabela discriminada do relatório de Paralisados (sempre inicia em página nova).
    // Retorna o número da página inicial (para o índice/bookmarks do PDF conjunto).
    function montarTabelaParalisados(doc, dados, comIndice) {
        const agora = new Date();
        const pw = doc.internal.pageSize.getWidth();
        const ph = doc.internal.pageSize.getHeight();
        const m = 12;
        const hoje = agora.toLocaleDateString('pt-BR');
        const hora = agora.toLocaleTimeString('pt-BR', { hour: '2-digit', minute: '2-digit' });

        // Do maior tempo paralisado para o menor (sem dados vai ao final)
        const ordenados = dados.slice().sort((a, b) => {
            const da = a.dias == null ? -Infinity : a.dias;
            const db = b.dias == null ? -Infinity : b.dias;
            return db - da;
        });

        doc.addPage();
        const paginaInicial = doc.internal.getNumberOfPages();
        tituloSecao(doc, m, m + 3, pw - 2 * m, 'Tabela discriminada — Processos Paralisados');

        const colunas = [
            { header: 'Processo', width: 30, get: (d) => d.processo },
            { header: 'Classe Processual', width: 56, get: (d) => d.classe },
            { header: 'Dias Paralisado', width: 20, get: (d) => (d.dias == null ? '' : String(d.dias)) },
            { header: 'Último Movimento', width: 70, get: (d) => d.ultimoMovimento },
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
            styles: { font: 'PublicSans', fontSize: 7.5, cellPadding: 1.6, textColor: COR.tintaSec,
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
            didDrawPage: () => desenharRodape(doc, TITULO_PARALISADOS, `${hoje} ${hora}`, pw, ph, m, comIndice),
        });

        return paginaInicial;
    }

    // ── PDF do relatório de Remessas em Aberto ──────────────────────────────────
    // Mesmo padrão do Paralisados, mas sem o KPI de "razão externa" (não existe nos dados
    // extraídos aqui) — linha 1 com 3 KPIs em vez de 4.

    const TITULO_REMESSAS = 'Remessas em Aberto';

    function gerarPDFRemessas(dados, somenteResumo) {
        const doc = novoDocPDF();
        montarResumoRemessas(doc, dados, true, false);
        doc.outline.add(null, 'Resumo', { pageNumber: 1 });
        if (!somenteResumo) {
            const pgTabela = montarTabelaRemessas(doc, dados, false);
            doc.outline.add(null, 'Tabela detalhada', { pageNumber: pgTabela });
        }
        const sufixo = somenteResumo ? '_resumo' : '';
        baixarBlob(doc.output('blob'), `remessas_abertas_projudi${sufixo}_${dataArquivo()}.pdf`);
    }

    // Página de RESUMO (KPIs + gráficos) do relatório de Remessas em Aberto, dentro de um
    // doc jsPDF já existente. ehPrimeiraSecao=false começa em página nova (uso no conjunto).
    function montarResumoRemessas(doc, dados, ehPrimeiraSecao, comIndice) {
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
        const maisParado = validos.slice().sort((a, b) => b.dias - a.dias)[0] || null;

        doc.setFont('PublicSans', 'bold'); doc.setFontSize(16); doc.setTextColor(...COR.tinta);
        doc.text(TITULO_REMESSAS, m, m + 2);
        doc.setFont('PublicSans', 'normal'); doc.setFontSize(9); doc.setTextColor(...COR.tintaSec);
        let subtituloRemessas = `Extraído em ${hoje} às ${hora}  •  ${dados.length} processo(s) em remessa`;
        const fraseCompRemessas = fraseCompetencias(dados);
        if (fraseCompRemessas) subtituloRemessas += `  •  ${fraseCompRemessas}`;
        const linhasSubtituloRemessas = doc.splitTextToSize(subtituloRemessas, uw);
        doc.text(linhasSubtituloRemessas, m, m + 8);
        const yLinhaRemessas = m + 8 + (linhasSubtituloRemessas.length - 1) * 4.2 + 3;
        doc.setDrawColor(...COR.azul); doc.setLineWidth(0.5); doc.line(m, yLinhaRemessas, pw - m, yLinhaRemessas);

        const gap = 6;
        // Linha 1: 3 KPIs centralizados — em remessa / tempo médio paralisado / prioritários
        const kY = yLinhaRemessas + 5;
        const kW3 = (uw - 2 * gap) / 3;
        const prioPct = validos.length ? Math.round(prioritarios.length / validos.length * 100) : 0;
        desenharCard(doc, m,                kY, kW3, 28, 'Processos em remessa', String(dados.length), [], true, COR.azul);
        desenharCard(doc, m + kW3 + gap,     kY, kW3, 28, 'Tempo médio paralisado', fmtDias(geral), [], true, COR.aqua);
        desenharCard(doc, m + 2*(kW3+gap),   kY, kW3, 28, 'Prioritários', String(prioritarios.length), [`${prioPct}% do total`], true, COR.vermelho);

        // Linha 2: tempo médio paralisado prioritários vs não prioritários (centralizados)
        const k2Y = kY + 28 + gap;
        const kW2 = (uw - gap) / 2;
        desenharCard(doc, m,           k2Y, kW2, 26, 'Tempo médio paralisado — Prioritários',     fmtDias(mediaPrio),    [`${prioritarios.length} processo(s)`], true, COR.vermelho);
        desenharCard(doc, m + kW2+gap, k2Y, kW2, 26, 'Tempo médio paralisado — Não prioritários', fmtDias(mediaNaoPrio), [`${naoPrioritarios.length} processo(s)`], true, COR.azul);

        // Card largo: processo paralisado há mais tempo (centralizado)
        const k3Y = k2Y + 26 + gap;
        let valMP = '—', subsMP = ['Nenhum registro com dados válidos'];
        if (maisParado) {
            valMP = `${maisParado.dias} dia${maisParado.dias === 1 ? '' : 's'}`;
            subsMP = [
                `Processo ${maisParado.processo}${maisParado.prioritario ? '  — PRIORITÁRIO' : ''}`,
                maisParado.classe || '',
            ];
        }
        desenharCard(doc, m, k3Y, uw, 26, 'Processo paralisado há mais tempo', valMP, subsMP, true, COR.vermelho);

        // Gráficos 1 e 2 (largura total, empilhados): dividem o espaço vertical restante da
        // página de forma que a soma das duas alturas + o espaçamento sempre caiba.
        const chartY = k3Y + 26 + gap + 2;
        const chartGap = 8;
        const disponivel = ph - m - chartY - 14; // reserva a faixa do rodapé
        const top10 = validos.slice().sort((a, b) => b.dias - a.dias).slice(0, 10)
            .map(d => ({ processo: d.processo, classe: d.classe, dias: d.dias, prioritario: d.prioritario }));
        const chart1Desejado = Math.max(30, top10.length * 8 + 8);
        const chart1H = Math.min(chart1Desejado, Math.max(30, disponivel - chartGap - 40));
        const chart2H = Math.max(30, disponivel - chart1H - chartGap);
        desenharTopDemorados(doc, m, chartY, uw, chart1H, 'Processos paralisados há mais tempo', top10);

        const chart2Y = chartY + chart1H + chartGap;
        const porClasse = agregarMedia(validos, 'classe', 'dias', 12);
        desenharBarras(doc, m, chart2Y, uw, chart2H, 'Tempo médio paralisado por Classe Processual', porClasse, fmtDias, COR.aqua);

        desenharRodape(doc, TITULO_REMESSAS, `${hoje} ${hora}`, pw, ph, m, comIndice);
    }

    // Tabela discriminada do relatório de Remessas em Aberto (sempre inicia em página
    // nova). Retorna o número da página inicial (para o índice/bookmarks do conjunto).
    function montarTabelaRemessas(doc, dados, comIndice) {
        const agora = new Date();
        const pw = doc.internal.pageSize.getWidth();
        const ph = doc.internal.pageSize.getHeight();
        const m = 12;
        const hoje = agora.toLocaleDateString('pt-BR');
        const hora = agora.toLocaleTimeString('pt-BR', { hour: '2-digit', minute: '2-digit' });

        // Do maior tempo paralisado para o menor (sem dados vai ao final)
        const ordenados = dados.slice().sort((a, b) => {
            const da = a.dias == null ? -Infinity : a.dias;
            const db = b.dias == null ? -Infinity : b.dias;
            return db - da;
        });

        doc.addPage();
        const paginaInicial = doc.internal.getNumberOfPages();
        tituloSecao(doc, m, m + 3, pw - 2 * m, 'Tabela discriminada — Remessas em Aberto');

        const colunas = [
            { header: 'Processo', width: 32, get: (d) => d.processo },
            { header: 'Classe Processual', width: 60, get: (d) => d.classe },
            { header: 'Dias Paralisado', width: 20, get: (d) => (d.dias == null ? '' : String(d.dias)) },
            { header: 'Último Movimento', width: 70, get: (d) => d.ultimoMovimento },
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
            styles: { font: 'PublicSans', fontSize: 7.5, cellPadding: 1.6, textColor: COR.tintaSec,
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
            didDrawPage: () => desenharRodape(doc, TITULO_REMESSAS, `${hoje} ${hora}`, pw, ph, m, comIndice),
        });

        return paginaInicial;
    }

    // ── Interface ───────────────────────────────────────────────────────────────

    function atualizarStatus(msg) {
        const el = document.getElementById('exportar-status');
        if (el) el.textContent = msg;
    }

    // "Limpar" nunca é desabilitado — é o botão de resgate caso a coleta trave (evita o
    // usuário ficar sem forma de apagar os dados presos e recomeçar).
    function desabilitarBotoes(desabilitar) {
        ['btn-coletar', 'btn-baixar', 'btn-pdf', 'btn-pdf-juiz'].forEach(id => {
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
        #btn-pdf, #btn-pdf-juiz { background-color: #34556b; border-color: #26404f; }
        #btn-limpar  { background-color: #8a3b3b; border-color: #6e2f2f; }
        #btn-preencher-pesquisar-tm { background-color: #1e6b1e; border-color: #145214; }
        .projudi-btn:disabled { background-color: #999; border-color: #777; cursor: not-allowed; }
        .projudi-select {
            margin-left: 6px; padding: 2px 4px; font-size: 0.85em;
            font-family: Verdana, Arial, sans-serif; border-radius: 3px; border: 1px solid #999;
        }
        #exportar-status { font-size: 0.8em; color: #555; margin-left: 8px; font-family: Verdana, Arial, sans-serif; }
        .projudi-chk-resumo {
            display: inline-flex; align-items: center; gap: 3px; margin-left: 8px;
            font-size: 0.8em; color: #444; font-family: Verdana, Arial, sans-serif; cursor: pointer;
        }
        .projudi-chk-resumo input[type="checkbox"] { margin: 0; }
    `);

    // Detecta qual relatório está na tela pelo cabeçalho da resultTable; se não houver
    // tabela reconhecível, tenta pela URL (útil na tela de busca da página de juntadas).
    function detectarConfig() {
        const thead = document.querySelector('table.resultTable thead');
        const cab = thead ? thead.textContent : '';
        let cfg = null;
        // CFG_SUSPENSOS vem antes de CFG_PARALISADOS: a tabela de Suspensos por Prazo
        // Indeterminado também tem a coluna "Dias Paralisado" (o regex de Paralisados
        // sozinho a reconheceria por engano), mas só ela tem "Início Suspensão".
        if (CFG_SUSPENSOS.detecta(cab)) cfg = CFG_SUSPENSOS;
        else if (CFG_TEMPOMEDIO.detecta(cab)) cfg = CFG_TEMPOMEDIO;
        else if (CFG_AUDIENCIAS.detecta(cab)) cfg = CFG_AUDIENCIAS;
        else if (CFG_PARALISADOS.detecta(cab)) cfg = CFG_PARALISADOS;
        else if (CFG_REMESSAS.detecta(cab)) cfg = CFG_REMESSAS;
        else if (CFG_JUNTADAS.detecta(cab)) cfg = CFG_JUNTADAS;
        else if (CFG_RETORNO.detecta(cab)) cfg = CFG_RETORNO;
        else if (CFG_CONCLUSOES.detecta(cab)) cfg = CFG_CONCLUSOES;
        else if (CFG_APREENSOES.detecta(cab)) cfg = CFG_APREENSOES;
        else if (/analisarJuntada\.do/i.test(location.pathname + location.search)) cfg = CFG_JUNTADAS;
        else if (/processoBuscaSuspenso\.do/i.test(location.pathname + location.search)) cfg = CFG_SUSPENSOS;
        else if (/processoBuscaParalisado\.do/i.test(location.pathname + location.search)) {
            cfg = opcaoBuscaParalisadoSelecionada() === '3' ? CFG_REMESSAS : CFG_PARALISADOS;
        }
        console.log(`[Projudi] detectarConfig — url=${location.pathname} thead=${!!thead} situacaoAudiencia=${situacaoAudienciaSelecionada()} cfg=${cfg ? cfg.prefixo : 'null'} cab="${cab.slice(0,80).replace(/\s+/g,' ')}"`);
        return cfg;
    }

    // Tela de filtros do relatório "Estatísticas de Conclusões" (tempo médio), antes da pesquisa.
    function formularioTempoMedio() {
        const form = document.getElementById('estatisticaConclusaoForm');
        return form && form.querySelector('input[name="situacao"]') ? form : null;
    }

    // Quantidade de meses completos buscados pelo relatório de Tempo Médio. Um período
    // grande (ex.: 1 ano de uma vez) deixa a pesquisa do Projudi lenta — em vez de um único
    // intervalo, cada opção dispara N pesquisas separadas, uma por mês completo (ver
    // mesesCompletos/prepararFilaMesesTempoMedio), acumulando tudo no mesmo relatório.
    const PERIODOS_TEMPOMEDIO = [
        { id: '1m',  rotulo: 'Último mês completo',          meses: 1 },
        { id: '6m',  rotulo: 'Últimos 6 meses completos',    meses: 6 },
        { id: '1a',  rotulo: 'Últimos 12 meses completos',   meses: 12 },
        { id: '2a',  rotulo: 'Últimos 24 meses completos',   meses: 24 },
    ];

    function formatarDataBR(d) {
        const dd = String(d.getDate()).padStart(2, '0');
        const mm = String(d.getMonth() + 1).padStart(2, '0');
        return `${dd}/${mm}/${d.getFullYear()}`;
    }

    // Gera os últimos "n" meses completos — o mês atual, ainda em andamento, NUNCA entra —,
    // do mais recente para o mais antigo. Ex.: hoje 24/08/2026, n=6 -> julho, junho, maio,
    // abril, março, fevereiro/2026, nessa ordem.
    function mesesCompletos(n, agora) {
        const base = agora || new Date();
        const meses = [];
        for (let i = 1; i <= n; i++) {
            const ini = new Date(base.getFullYear(), base.getMonth() - i, 1);
            const fim = new Date(base.getFullYear(), base.getMonth() - i + 1, 0); // dia 0 = último dia do mês anterior
            meses.push({
                ini: formatarDataBR(ini),
                fim: formatarDataBR(fim),
                rotulo: `${String(ini.getMonth() + 1).padStart(2, '0')}/${ini.getFullYear()}`,
            });
        }
        return meses;
    }

    const CHAVE_FILA_MESES_TM = 'projudi_tempomedio_fila_meses';

    // Monta (ou reinicia) a fila de meses a pesquisar para o Tempo Médio, a partir do
    // período escolhido — chamado uma única vez, no início da extração (automação ou botão
    // manual), nunca a cada mês (senão a fila voltaria sempre ao tamanho cheio).
    function prepararFilaMesesTempoMedio(periodoId) {
        const periodo = PERIODOS_TEMPOMEDIO.find(p => p.id === periodoId) || PERIODOS_TEMPOMEDIO.find(p => p.id === '1m');
        const fila = mesesCompletos(periodo.meses);
        store.setItem(CHAVE_FILA_MESES_TM, JSON.stringify(fila));
        // Limpa o período acumulado de uma rodada anterior (ver preencherEPesquisarTempoMedio)
        // para não herdar por engano o "fim" de uma extração antiga.
        store.removeItem('projudi_tempomedio_periodo');
        return fila;
    }

    function lerFilaMesesTempoMedio() {
        // Mesma proteção usada em lerDadosDe/desembrulharArray: o valor às vezes volta
        // JSON-codificado em camadas (visto em outros relatórios) — sem isso, um
        // JSON.parse único deixava "fila" como STRING (não array), e fila[0] virava um
        // caractere solto em vez do objeto do mês, gerando datas "undefined" na pesquisa.
        return desembrulharArray(store.getItem(CHAVE_FILA_MESES_TM)) || [];
    }

    // Marca Situação=Analisadas, Tipo=Analítico, define a data inicial/final com o PRÓXIMO
    // mês da fila (ver prepararFilaMesesTempoMedio — nunca busca mais de um mês de cada
    // vez) e clica em Pesquisar. O diagnóstico anterior (3s) apontava falha, mas era falso
    // negativo: o site do Projudi é lento e a navegação real ainda estava em andamento
    // quando o diagnóstico rodou — por isso o timeout de checagem agora é bem mais longo.
    function preencherEPesquisarTempoMedio() {
        const form = formularioTempoMedio();
        if (!form) return;

        const fila = lerFilaMesesTempoMedio();
        const mes = fila[0];
        if (!mes) { console.warn('[Projudi TM] preencherEPesquisarTempoMedio chamado sem fila de meses — nada a pesquisar'); return; }
        // Remove o mês do início da fila agora — quando essa pesquisa terminar de coletar
        // (ver criarColetor/continuar), o próximo item (se houver) dispara uma nova rodada.
        store.setItem(CHAVE_FILA_MESES_TM, JSON.stringify(fila.slice(1)));

        const radioAnalisadas = form.querySelector('input[name="situacao"][value="A"]');
        const radioAnalitico = form.querySelector('input[name="analitico"][value="true"]');
        console.log(`[Projudi TM] radioAnalisadas encontrado=${!!radioAnalisadas} radioAnalitico encontrado=${!!radioAnalitico} mês="${mes.rotulo}" (${fila.length - 1} restante(s) na fila)`);
        if (radioAnalisadas) radioAnalisadas.checked = true;
        if (radioAnalitico) radioAnalitico.checked = true;

        const campoInicio = form.querySelector('input[name="dataInicio"]');
        if (campoInicio) {
            // Campos de data vêm com "disabled" no HTML original: se ficarem desabilitados,
            // o navegador NÃO os envia no submit. Precisamos habilitar antes de definir o valor.
            campoInicio.disabled = false;
            campoInicio.value = mes.ini;
        }
        const campoFim = form.querySelector('input[name="dataFim"]');
        if (campoFim) { campoFim.disabled = false; campoFim.value = mes.fim; } // idem — sem isso a busca não teria fim de mês
        console.log(`[Projudi TM] campos preenchidos — dataInicio="${campoInicio ? campoInicio.value : '?'}" dataFim="${campoFim ? campoFim.value : '?'}"`);

        // Salva o período ACUMULADO (não só o deste mês) para uso posterior no PDF — como a
        // fila roda do mês mais recente para o mais antigo, o "fim" só é gravado na primeira
        // pesquisa (mês mais recente) e o "ini" vai sendo sobrescrito a cada mês seguinte,
        // terminando no início do mês mais antigo depois que a fila inteira for processada.
        const periodoAnterior = (() => {
            // Mesma proteção contra JSON codificado em camadas usada em lerFilaMesesTempoMedio.
            let v = store.getItem('projudi_tempomedio_periodo') || '{}';
            let t = 0;
            while (typeof v === 'string' && t < 5) { try { v = JSON.parse(v); } catch (e) { return {}; } t++; }
            return (v && typeof v === 'object' && !Array.isArray(v)) ? v : {};
        })();
        store.setItem('projudi_tempomedio_periodo', JSON.stringify({
            ini: campoInicio ? campoInicio.value : '',
            fim: periodoAnterior.fim || (campoFim ? campoFim.value : ''),
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

    // Tela de filtros de Processos Paralisados/Remessas (processoBuscaParalisado.do) —
    // as duas páginas de resultado são a MESMA tela, só muda o rádio "opcaoBusca"
    // marcado (ver opcaoBuscaParalisadoSelecionada) e o "Mínimo de dias paralisado".
    function formularioParalisado() {
        const form = document.getElementById('processoBuscaParalisadoForm');
        return form && form.querySelector('input[name="opcaoBusca"]') ? form : null;
    }

    // opcaoBuscaValor: '1' = Na secretaria (Paralisados), '3' = Em remessa, exceto
    // processos conclusos (Remessas em Aberto). chaveParaAutoIniciar, se informada, marca
    // a extração para começar sozinha assim que os resultados aparecerem.
    function preencherEPesquisarParalisado(opcaoBuscaValor, diasMinimo, chaveParaAutoIniciar) {
        const form = formularioParalisado();
        if (!form) return;

        const radio = form.querySelector(`input[name="opcaoBusca"][value="${opcaoBuscaValor}"]`);
        console.log(`[Projudi Paralisado] radio opcaoBusca=${opcaoBuscaValor} encontrado=${!!radio}`);
        if (radio) {
            radio.checked = true;
            radio.dispatchEvent(new Event('click', { bubbles: true }));
            radio.dispatchEvent(new Event('change', { bubbles: true }));
        }

        const campoDias = form.querySelector('input[name="numeroDiasParalisados"]');
        if (campoDias) {
            campoDias.value = String(diasMinimo);
            campoDias.dispatchEvent(new Event('input', { bubbles: true }));
            campoDias.dispatchEvent(new Event('change', { bubbles: true }));
        }
        console.log(`[Projudi Paralisado] campos preenchidos — opcaoBusca=${opcaoBuscaValor} diasMinimo="${campoDias ? campoDias.value : '?'}"`);

        if (chaveParaAutoIniciar) store.setItem('projudi_paralisado_auto_iniciar', chaveParaAutoIniciar);

        const btn = document.getElementById('pesquisar') || form.querySelector('input[type="submit"]');
        console.log(`[Projudi Paralisado] botão de pesquisa encontrado=${!!btn}; clicando em 1,5s`);

        setTimeout(() => {
            console.log('[Projudi Paralisado] clicando em Pesquisar — o site pode demorar para responder, aguarde.');
            if (btn && !btn.disabled) btn.click(); else form.submit();

            // Diagnóstico tardio (site é lento; não dispara nenhum reenvio, só informa).
            setTimeout(() => {
                const aindaNoFormulario = !!document.getElementById('processoBuscaParalisadoForm');
                const temResultado = !!document.querySelector('table.resultTable tbody tr');
                console.log(`[Projudi Paralisado] diagnóstico 15s depois — aindaNoFormulario=${aindaNoFormulario} temResultado=${temResultado}`);
                if (aindaNoFormulario && !temResultado) {
                    console.warn('[Projudi Paralisado] ainda sem resultado após 15s — o site pode estar lento; se persistir, clique em Pesquisar manualmente.');
                }
            }, 15000);
        }, 1500);
    }

    // Tela de "Listagem" de Audiências (audiencia/busca.do) — igual Paralisados/Remessas,
    // o formulário de filtros e a table.resultTable (com o que quer que tenha sido
    // pesquisado por último, ex.: "Para Hoje") convivem na MESMA página desde o primeiro
    // carregamento.
    function formularioAudiencias() {
        const form = document.getElementById('audienciaForm');
        return form && form.querySelector('input[name="idSituacaoAudiencia"]') ? form : null;
    }

    // Marca a situação "Pendentes" (idSituacaoAudiencia=8) e clica em Pesquisar.
    function preencherEPesquisarAudiencias() {
        const form = formularioAudiencias();
        if (!form) return;

        const radio = form.querySelector('input[name="idSituacaoAudiencia"][value="8"]');
        console.log(`[Projudi Audiências] radio Pendentes encontrado=${!!radio}`);
        if (radio) {
            radio.checked = true;
            radio.dispatchEvent(new Event('click', { bubbles: true }));
            radio.dispatchEvent(new Event('change', { bubbles: true }));
        }

        const btn = document.getElementById('pesquisar') || form.querySelector('input[type="submit"]');
        console.log(`[Projudi Audiências] botão de pesquisa encontrado=${!!btn}; clicando em 1,5s`);

        setTimeout(() => {
            console.log('[Projudi Audiências] clicando em Pesquisar — o site pode demorar para responder, aguarde.');
            if (btn && !btn.disabled) btn.click(); else form.submit();

            // Diagnóstico tardio (site é lento; não dispara nenhum reenvio, só informa).
            setTimeout(() => {
                const aindaSemPendentes = detectarConfig() !== CFG_AUDIENCIAS;
                console.log(`[Projudi Audiências] diagnóstico 15s depois — aindaSemPendentes=${aindaSemPendentes}`);
                if (aindaSemPendentes) {
                    console.warn('[Projudi Audiências] ainda sem resultado de "Pendentes" após 15s — o site pode estar lento; se persistir, clique em Pesquisar manualmente.');
                }
            }, 15000);
        }, 1500);
    }

    // URL esperada da tela de resultados de cada relatório — usada só para confirmar que
    // estamos mesmo na página certa antes de considerar "0 registros" (ver injetarBotoes).
    function urlEsperadaRelatorio(navAlvo) {
        if (navAlvo === 'juntadas') return /analisarJuntada\.do/i;
        if (navAlvo === 'retorno' || navAlvo === 'conclusoes') return /conclusao\.do/i;
        if (navAlvo === 'tempomedio') return /conclusao\/estatistica\.do/i;
        if (navAlvo === 'paralisados' || navAlvo === 'remessas') return /processoBuscaParalisado\.do/i;
        if (navAlvo === 'suspensos') return /processoBuscaSuspenso\.do/i;
        if (navAlvo === 'audiencias') return /audiencia\/busca\.do/i;
        if (navAlvo === 'audienciasdesignadas') return /audiencia\/pautaAudiencia\.do/i;
        if (navAlvo === 'audienciasrealizadas') return /audiencia\/estatistica\.do/i;
        if (navAlvo === 'apreensoes') return /processo\/criminal\/apreensao\.do/i;
        return null;
    }

    function injetarBotoes() {
        console.log(`[Projudi] injetarBotoes — url=${location.pathname} estadoAuto=${store.getItem(AUTO_ESTADO)}`);
        const buttonBar = document.querySelector('table.buttonBar td.buttons');
        if (!buttonBar) {
            // Quando uma busca não encontra NENHUM registro, algumas telas do Projudi
            // (Juntadas/Retorno/Conclusões) nem chegam a renderizar a tabela de
            // resultados nem a buttonBar — só um aviso de "nenhum resultado". Sem esse
            // tratamento, a automação ficava parada para sempre em "coletando_X" nessa
            // tela (nunca chamava coletor.iniciar()), e ao trocar de atuação manualmente
            // o usuário podia achar que travou e clicar em "Limpar", apagando os dados já
            // acumulados de atuações anteriores — o que parecia "o relatório ignorou os
            // dados da atuação anterior". Agora, se a automação está esperando coletar
            // exatamente este relatório e a URL bate com a esperada, tratamos como "0
            // registros" e avançamos para o próximo item da fila, sem perder nada do que
            // já foi coletado (esse relatório só entra 0 no PDF conjunto).
            const estadoAuto = store.getItem(AUTO_ESTADO);
            if (estadoAuto && estadoAuto.startsWith('coletando_')) {
                const rel = relatorioPorChave(estadoAuto.slice('coletando_'.length));
                const urlRe = rel && urlEsperadaRelatorio(rel.navAlvo);
                if (rel && urlRe && urlRe.test(location.href)) {
                    console.log(`[Auto Projudi] "${rel.rotulo}" sem tabela de resultados (0 registros) — avançando automação`);
                    // Sem isso, "0 registros" (nunca chega a passar por criarColetor) ficava
                    // indistinguível de "nunca coletado" — e o card sumia do PDF conjunto em
                    // vez de mostrar 0 pendências (ver KEY_COLETADO em criarColetor).
                    store.setItem(rel.cfg.prefixo + 'coletado', '1');
                    avancarAutomacao(rel.cfg);
                }
            }
            return;
        }

        // Tela de filtros do relatório de Tempo Médio (ainda sem resultados): select do
        // período + botão de preencher+pesquisar; os botões de coleta/exportação fazem
        // sentido depois da busca.
        if (formularioTempoMedio() && !document.querySelector('table.resultTable')) {
            // Automação: se chegamos aqui vindos do fluxo de automação, preenche e pesquisa
            // sozinho (sem esperar clique manual) — a fila de meses já foi preparada em
            // iniciarAutomacao() (ou por uma rodada anterior, ver criarColetor/continuar).
            if (store.getItem(AUTO_ESTADO) === 'preenchendo_tempomedio') {
                console.log('[Projudi TM] automação: preenchendo e pesquisando o próximo mês da fila');
                store.setItem(AUTO_ESTADO, 'coletando_tempomedio');
                preencherEPesquisarTempoMedio();
                return;
            }

            const sel = document.createElement('select');
            sel.id = 'sel-periodo-tm';
            sel.className = 'projudi-select';
            sel.title = 'Quantos meses completos buscar (sempre em pesquisas separadas por mês)';
            PERIODOS_TEMPOMEDIO.forEach(p => {
                const opt = document.createElement('option');
                opt.value = p.id;
                opt.textContent = p.rotulo;
                sel.appendChild(opt);
            });
            sel.value = '1m'; // padrão: último mês completo
            buttonBar.appendChild(sel);

            const b = document.createElement('button');
            b.id = 'btn-preencher-pesquisar-tm';
            b.type = 'button';
            b.className = 'projudi-btn';
            b.title = 'Marca Analisadas + Analítico e pesquisa mês a mês, do mais recente ao mais antigo (o site pode demorar a responder)';
            b.textContent = 'Preencher e Pesquisar';
            b.onclick = () => {
                // Fora da automação (uso avulso, direto nesta página): monta a fila de
                // meses e conduz o mesmo fluxo de "coletando_tempomedio" usado pela
                // automação — é o que permite avançar sozinho de mês em mês a cada
                // recarregamento de página (ver criarColetor/continuar).
                prepararFilaMesesTempoMedio(sel.value);
                store.setItem('projudi_auto_fila', JSON.stringify(['tempomedio']));
                store.setItem(AUTO_ESTADO, 'coletando_tempomedio');
                store.setItem('projudi_auto_lock', String(Date.now()));
                preencherEPesquisarTempoMedio();
            };
            buttonBar.appendChild(b);
            return;
        }

        // Tela de filtros de Processos Paralisados/Remessas (mesma página para os dois
        // relatórios — só muda o rádio "opcaoBusca" marcado e o mínimo de dias). Diferente
        // do Tempo Médio, aqui o form e a table.resultTable (com cabeçalho, mesmo sem
        // linhas) convivem na MESMA página desde o primeiro carregamento — por isso a
        // automação precisa decidir com base no ESTADO (preenchendo_X), e não na presença
        // de table.resultTable, senão nunca preenche nem pesquisa (fica "parado").
        if (formularioParalisado()) {
            const estadoAtual = store.getItem(AUTO_ESTADO);
            if (estadoAtual === 'preenchendo_paralisados' || estadoAtual === 'preenchendo_remessas') {
                const chave = estadoAtual.slice('preenchendo_'.length);
                const opcaoBuscaValor = chave === 'remessas' ? '3' : '1';
                console.log(`[Projudi Paralisado] automação: preenchendo e pesquisando (${chave})`);
                store.setItem(AUTO_ESTADO, 'coletando_' + chave);
                preencherEPesquisarParalisado(opcaoBuscaValor, 30, chave);
                return;
            }

            // Uso manual (fora da automação), só antes de qualquer pesquisa ter sido feita:
            // a página é compartilhada pelos dois relatórios, então oferecemos um botão
            // para cada um. Uma vez que já há resultados na tela, cai para o fluxo normal
            // de coleta/exportação mais abaixo.
            if (!document.querySelector('table.resultTable tbody tr')) {
                const bParalisados = document.createElement('button');
                bParalisados.type = 'button';
                bParalisados.className = 'projudi-btn';
                bParalisados.title = 'Marca "Na secretaria", define 30 dias mínimos e pesquisa (o site pode demorar a responder)';
                bParalisados.textContent = 'Preencher e Pesquisar (Paralisados)';
                bParalisados.onclick = () => preencherEPesquisarParalisado('1', 30, 'paralisados');
                buttonBar.appendChild(bParalisados);

                const bRemessas = document.createElement('button');
                bRemessas.type = 'button';
                bRemessas.className = 'projudi-btn';
                bRemessas.title = 'Marca "Em remessa, exceto processos conclusos", define 30 dias mínimos e pesquisa (o site pode demorar a responder)';
                bRemessas.textContent = 'Preencher e Pesquisar (Remessas)';
                bRemessas.onclick = () => preencherEPesquisarParalisado('3', 30, 'remessas');
                buttonBar.appendChild(bRemessas);
                return;
            }
        }

        // Tela de "Listagem" de Audiências: mesmo padrão de Paralisados/Remessas (form +
        // resultTable na mesma página desde o início) — mas aqui não dá pra usar "ainda
        // sem linha nenhuma" pra decidir se falta pesquisar, porque a tela já carrega com
        // resultados de OUTRO filtro (ex.: "Para Hoje", marcado por padrão). Em vez disso,
        // confere se o que já está na tela é mesmo "Pendentes" (detectarConfig confere
        // cabeçalho + rádio marcado).
        if (formularioAudiencias()) {
            const estadoAtual = store.getItem(AUTO_ESTADO);
            if (estadoAtual === 'preenchendo_audiencias') {
                console.log('[Projudi Audiências] automação: preenchendo e pesquisando');
                store.setItem(AUTO_ESTADO, 'coletando_audiencias');
                preencherEPesquisarAudiencias();
                return;
            }
            if (detectarConfig() !== CFG_AUDIENCIAS) {
                const bAudiencias = document.createElement('button');
                bAudiencias.type = 'button';
                bAudiencias.className = 'projudi-btn';
                bAudiencias.title = 'Marca "Pendentes" e pesquisa (o site pode demorar a responder)';
                bAudiencias.textContent = 'Preencher e Pesquisar (Audiências Pendentes)';
                bAudiencias.onclick = () => preencherEPesquisarAudiencias();
                buttonBar.appendChild(bAudiencias);
                return;
            }
        }

        // Tela "Ver Pauta de Horários" — não passa pelo coletor genérico (a tabela é
        // aninhada por dia/local, não uma table.resultTable simples de linhas soltas — ver
        // coletarAudienciasDesignadas), e não tem paginação (o Projudi devolve tudo numa
        // página só, mesmo pedindo 10 anos).
        if (formularioPautaAudiencias()) {
            const estadoAtual = store.getItem(AUTO_ESTADO);
            if (estadoAtual === 'preenchendo_audienciasdesignadas') {
                console.log('[Projudi Audiências Designadas] automação: preenchendo e pesquisando');
                store.setItem(AUTO_ESTADO, 'coletando_audienciasdesignadas');
                preencherEPesquisarPautaAudiencias();
                return;
            }
            if (estadoAtual === 'coletando_audienciasdesignadas' && pautaAudienciasTemResultados()) {
                // O clique em Pesquisar recarregou esta mesma tela já com os resultados —
                // dispara a extração (assíncrona: expande as linhas do último dia).
                coletarAudienciasDesignadas();
                return;
            }
            if (!pautaAudienciasTemResultados()) {
                const bPreencher = document.createElement('button');
                bPreencher.type = 'button';
                bPreencher.className = 'projudi-btn';
                bPreencher.title = 'Marca todos os tipos, define Data Fim em 10 anos e pesquisa (o site pode demorar a responder)';
                bPreencher.textContent = 'Preencher e Pesquisar (Audiências Designadas)';
                bPreencher.onclick = () => preencherEPesquisarPautaAudiencias();
                buttonBar.appendChild(bPreencher);
                return;
            }
            const bExtrair = document.createElement('button');
            bExtrair.type = 'button';
            bExtrair.className = 'projudi-btn';
            bExtrair.title = 'Lê a pauta exibida e gera o resumo (total, último dia, processos daquele dia, última data por tipo)';
            bExtrair.textContent = 'Extrair Audiências Designadas';
            bExtrair.onclick = () => coletarAudienciasDesignadas();
            buttonBar.appendChild(bExtrair);
            return;
        }

        // Tela "Estatísticas de Audiência" (audiencia/estatistica.do) — form + resultado
        // (só totais, sem tabela discriminada) na mesma página. A automação preenche o
        // período e dispara uma pesquisa "geral", depois uma pesquisa por usuário (fila) —
        // ver iniciarBuscaAudienciasRealizadas/processarResultadoAudienciasRealizadas.
        if (formularioAudienciasRealizadas()) {
            const estadoAtual = store.getItem(AUTO_ESTADO);
            const temResultado = !!document.querySelector('table.resultTable');
            if (estadoAtual === 'preenchendo_audienciasrealizadas' || estadoAtual === 'coletando_audienciasrealizadas') {
                if (temResultado) processarResultadoAudienciasRealizadas();
                else iniciarBuscaAudienciasRealizadas();
                return;
            }
            if (!temResultado) {
                const bPreencher = document.createElement('button');
                bPreencher.type = 'button';
                bPreencher.className = 'projudi-btn';
                bPreencher.title = 'Preenche o período (últimos 3 anos até o mês anterior) e pesquisa o total geral (o site pode demorar a responder)';
                bPreencher.textContent = 'Preencher e Pesquisar (Audiências Realizadas)';
                bPreencher.onclick = () => {
                    store.setItem(AUTO_ESTADO, 'preenchendo_audienciasrealizadas');
                    iniciarBuscaAudienciasRealizadas();
                };
                buttonBar.appendChild(bPreencher);
                return;
            }
            const bExtrair = document.createElement('button');
            bExtrair.type = 'button';
            bExtrair.className = 'projudi-btn';
            bExtrair.title = 'Lê o total exibido e, em seguida, pesquisa usuário por usuário até completar o detalhamento (pode demorar bastante)';
            bExtrair.textContent = 'Extrair Audiências Realizadas';
            bExtrair.onclick = () => {
                store.setItem(AUTO_ESTADO, 'coletando_audienciasrealizadas');
                processarResultadoAudienciasRealizadas();
            };
            buttonBar.appendChild(bExtrair);
            return;
        }

        // Tela de Apreensões (processo/criminal/apreensao.do) — form + table.resultTable
        // na MESMA página desde o primeiro carregamento (mesmo padrão de Paralisados/
        // Audiências). A automação decide pelo ESTADO (preenchendo_apreensoes), não pela
        // presença de resultados — a página já carrega com os filtros padrão (pendentes)
        // aplicados, então "já tem tabela" não distingue "ainda não pesquisei nesta
        // rodada" de "resultado de uma pesquisa anterior".
        if (formularioApreensoes()) {
            const estadoAtual = store.getItem(AUTO_ESTADO);
            if (estadoAtual === 'preenchendo_apreensoes') {
                console.log('[Projudi Apreensões] automação: preenchendo e pesquisando');
                store.setItem(AUTO_ESTADO, 'coletando_apreensoes');
                preencherEPesquisarApreensoes();
                return;
            }
            // Uso manual (fora da automação): a tela já convive com resultados de uma
            // pesquisa anterior (própria ou de outra sessão), então não há como usar
            // "sem linha nenhuma" para decidir se o botão deve aparecer.
            if (estadoAtual !== 'coletando_apreensoes') {
                const bPendentes = document.createElement('button');
                bPendentes.type = 'button';
                bPendentes.className = 'projudi-btn';
                bPendentes.title = 'Pesquisa com os filtros padrão da tela (Apreensão não encerrada) — não altera nenhum campo';
                bPendentes.textContent = 'Preencher e Pesquisar (Apreensões Pendentes)';
                bPendentes.onclick = () => preencherEPesquisarApreensoes();
                buttonBar.appendChild(bPendentes);
            }
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

        // Os botões de extração individual ficam ocultos por padrão (ver
        // CHAVE_MOSTRAR_BOTOES) — a extração automatizada não depende deles, chama
        // coletor.iniciar()/pdf()/etc. diretamente. O usuário reabilita pelo painel.
        const wrapExtracao = document.createElement('span');
        wrapExtracao.id = 'projudi-extracao-individual';
        if (!mostrarBotoesIndividuais()) wrapExtracao.style.display = 'none';
        buttonBar.appendChild(wrapExtracao);

        wrapExtracao.appendChild(mk('btn-coletar', 'Percorre todas as páginas e acrescenta aos dados já coletados', () => coletor.iniciar()));
        wrapExtracao.appendChild(mk('btn-baixar', 'Junta tudo o que foi coletado e baixa a planilha Excel', () => coletor.baixar()));
        // Conclusões só oferece planilha ou PDF por Juiz (abaixo) — o botão genérico de
        // PDF (resumo único + tabela, sem distinguir por juiz) não faz sentido aqui e
        // ficaria redundante ao lado do PDF por Juiz.
        if ((cfg.pdf || cfg.pdfCustom) && !cfg.pdfPorJuiz) {
            wrapExtracao.appendChild(mk('btn-pdf', 'Gera um PDF com painel, gráficos e a tabela completa', () => {
                const chk = document.getElementById('chk-somente-resumo');
                coletor.pdf(!!(chk && chk.checked));
            }));
            const rotuloResumo = document.createElement('label');
            rotuloResumo.className = 'projudi-chk-resumo';
            rotuloResumo.title = 'Gera o PDF só com o resumo (KPIs e gráficos), sem a tabela discriminada de processos';
            rotuloResumo.innerHTML = '<input type="checkbox" id="chk-somente-resumo"> Só resumo (sem tabela)';
            wrapExtracao.appendChild(rotuloResumo);
        }
        if (cfg.pdfPorJuiz) {
            wrapExtracao.appendChild(mk('btn-pdf-juiz',
                'Gera um PDF com uma seção por juiz responsável (pendências por tipo, agrupador e pré-análise) — pronto para expedir a cada um',
                () => coletor.pdfPorJuiz(), 'PDF por Juiz'));
        }
        wrapExtracao.appendChild(mk('btn-limpar', 'Apaga os dados acumulados deste relatório', () => coletor.limpar(), 'Limpar'));

        const status = document.createElement('span');
        status.id = 'exportar-status';
        wrapExtracao.appendChild(status);

        const estadoAuto = store.getItem(AUTO_ESTADO);
        const relAtual = relatorioPorCfg(cfg);
        const querColetarAuto = !!relAtual && estadoAuto === 'coletando_' + relAtual.key && cfg !== CFG_TEMPOMEDIO;
        const autoIniciarTM = cfg === CFG_TEMPOMEDIO && store.getItem('projudi_tempomedio_auto_iniciar') === '1';
        const chaveAutoIniciarParalisado = store.getItem('projudi_paralisado_auto_iniciar');
        const autoIniciarParalisado = !!chaveAutoIniciarParalisado && !!relAtual && relAtual.key === chaveAutoIniciarParalisado;

        console.log(`[Projudi] injetarBotoes — cfg=${cfg.prefixo} rodando=${coletor.rodando()} obsoleta=${coletor.obsoleta()} querColetarAuto=${querColetarAuto} autoIniciarTM=${autoIniciarTM} autoIniciarParalisado=${autoIniciarParalisado}`);

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
        } else if (autoIniciarParalisado) {
            store.removeItem('projudi_paralisado_auto_iniciar');
            console.log(`[Projudi Paralisado] flag auto_iniciar detectada (${chaveAutoIniciarParalisado}) — iniciando extração automaticamente`);
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
    // Marca o início/fim de uma rodada de automação (ver iniciarAutomacao/passoAutomacao)
    // — usados só para mostrar o tempo decorrido no painel (ver atualizarPainel).
    const CHAVE_AUTO_INICIO = 'projudi_auto_inicio';
    const CHAVE_AUTO_FIM = 'projudi_auto_fim';

    // "Xh Ym Zs" / "Ym Zs" / "Zs", sempre com o menor número de unidades necessário.
    function formatarDuracao(ms) {
        const totalSeg = Math.max(0, Math.round(ms / 1000));
        const h = Math.floor(totalSeg / 3600);
        const m = Math.floor((totalSeg % 3600) / 60);
        const s = totalSeg % 60;
        if (h > 0) return `${h}h ${m}m ${s}s`;
        if (m > 0) return `${m}m ${s}s`;
        return `${s}s`;
    }

    // Botões de extração individual (Extrair/Baixar/PDF/Limpar) em cada tela de relatório
    // ficam OCULTOS por padrão — a extração normal é pelo painel de automação da página
    // inicial. O usuário reabilita marcando a opção no próprio painel (ver injetarPainel).
    const CHAVE_MOSTRAR_BOTOES = 'projudi_mostrar_botoes_individuais';
    function mostrarBotoesIndividuais() { return store.getItem(CHAVE_MOSTRAR_BOTOES) === '1'; }

    function desembrulharArray(valor) {
        let v = valor, t = 0;
        while (typeof v === 'string' && t < 5) { try { v = JSON.parse(v); } catch (e) { break; } t++; }
        return Array.isArray(v) ? v : null;
    }

    // Mesma proteção que desembrulharArray, mas para um objeto plano (não array) — usado
    // para valores como o período do Tempo Médio, que também podem voltar do localStorage
    // JSON-codificados em camadas.
    function desembrulharObjeto(valor) {
        let v = valor, t = 0;
        while (typeof v === 'string' && t < 5) { try { v = JSON.parse(v); } catch (e) { return null; } t++; }
        return (v && typeof v === 'object' && !Array.isArray(v)) ? v : null;
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
    // Ordem pedida pelo usuário: Juntadas, Retorno de Conclusão, Processos Paralisados,
    // Remessas em Aberto e, por último, Conclusões pendentes.
    // "dominio" agrupa os checkboxes do painel de automação (ver injetarPainel), espelhando
    // a mesma divisão Cartório/Gabinete usada no PDF conjunto; "curto" é o rótulo compacto
    // usado na grade de contagens.
    const REPORTS_AUTOMACAO = [
        // O link de Suspensos por Prazo Indeterminado abre a tabela de resultados
        // direto (sem tela de filtros), igual Juntadas/Retorno/Conclusões.
        { key: 'suspensos',   cfg: CFG_SUSPENSOS,   navAlvo: 'suspensos',   rotulo: 'Suspensos p/ Prazo Indeterminado', curto: 'Suspensos',    dominio: 'cartorio', precisaPreencher: false },
        { key: 'juntadas',    cfg: CFG_JUNTADAS,    navAlvo: 'juntadas',    rotulo: 'Juntadas',              curto: 'Juntadas',    dominio: 'cartorio', precisaPreencher: false },
        { key: 'retorno',     cfg: CFG_RETORNO,     navAlvo: 'retorno',     rotulo: 'Retorno de Conclusos',   curto: 'Retorno',     dominio: 'cartorio', precisaPreencher: false },
        // Paralisados/Remessas caem na tela de filtros (com o mínimo de dias e o rádio de
        // situação), não direto nos resultados — por isso também precisam do passo de
        // preencher+pesquisar antes de coletar.
        { key: 'paralisados', cfg: CFG_PARALISADOS, navAlvo: 'paralisados', rotulo: 'Processos Paralisados',  curto: 'Paralisados', dominio: 'cartorio', precisaPreencher: true },
        { key: 'remessas',    cfg: CFG_REMESSAS,    navAlvo: 'remessas',    rotulo: 'Remessas em Aberto',     curto: 'Remessas',    dominio: 'cartorio', precisaPreencher: true },
        { key: 'conclusoes',  cfg: CFG_CONCLUSOES,  navAlvo: 'conclusoes',  rotulo: 'Conclusões',             curto: 'Conclusões',  dominio: 'gabinete', precisaPreencher: false },
        // Reabilitado (branch tempo-medio-teste) — busca em meses completos separados em
        // vez de um período único (ver mesesCompletos/prepararFilaMesesTempoMedio), para
        // evitar pesquisas grandes e lentas no Projudi. Vem DESMARCADO por padrão no
        // painel (ver injetarPainel): o usuário precisa marcar explicitamente para incluí-lo.
        { key: 'tempomedio',  cfg: CFG_TEMPOMEDIO,  navAlvo: 'tempomedio',  rotulo: 'Tempo Médio',            curto: 'Tempo Médio', dominio: 'cartorio', precisaPreencher: true },
        // Primeiro item específico da categoria Crime (ver CATEGORIAS_PAINEL/
        // categoriaEspecifica em injetarPainel) — não entra nos grupos Cartório/Gabinete
        // do Cível-Geral, só aparece na seção própria da aba Crime.
        { key: 'audiencias',  cfg: CFG_AUDIENCIAS,  navAlvo: 'audiencias',  rotulo: 'Audiências Pendentes',   curto: 'Audiências', categoriaEspecifica: 'crime', precisaPreencher: true },
        // Segundo item específico do Crime — resumo + tabela da Pauta de Horários (ver
        // coletarAudienciasDesignadas). Como cada extração recalcula tudo do zero (a
        // "página" gravada é sempre um único resumo), "Extrair mais" não faz sentido aqui.
        { key: 'audienciasdesignadas', cfg: CFG_AUDIENCIAS_DESIGNADAS, navAlvo: 'audienciasdesignadas', rotulo: 'Audiências Designadas', curto: 'Aud. Designadas', categoriaEspecifica: 'crime', precisaPreencher: true },
        // Terceiro item específico do Crime — totais de Estatísticas de Audiência, geral e
        // por usuário (ver iniciarBuscaAudienciasRealizadas/finalizarAudienciasRealizadas).
        // Assim como Audiências Designadas, cada extração recalcula tudo do zero.
        { key: 'audienciasrealizadas', cfg: CFG_AUDIENCIAS_REALIZADAS, navAlvo: 'audienciasrealizadas', rotulo: 'Audiências Realizadas', curto: 'Aud. Realizadas', categoriaEspecifica: 'crime', precisaPreencher: true },
        // Quarto item específico do Crime — apreensões pendentes; internamente roda em
        { key: 'apreensoes', cfg: CFG_APREENSOES, navAlvo: 'apreensoes', rotulo: 'Apreensões Pendentes', curto: 'Apreensões', categoriaEspecifica: 'crime', precisaPreencher: true },
    ];
    const GRUPOS_AUTOMACAO = [
        { chave: 'cartorio', rotulo: 'Cartório' },
        { chave: 'gabinete', rotulo: 'Gabinete' },
    ];
    function relatorioPorChave(key) { return REPORTS_AUTOMACAO.find(r => r.key === key); }
    function relatorioPorCfg(cfg) { return REPORTS_AUTOMACAO.find(r => r.cfg === cfg); }

    function lerFilaAutomacao() {
        try { return JSON.parse(store.getItem('projudi_auto_fila') || '[]'); } catch (e) { return []; }
    }

    // Sobe por window.parent enquanto for a mesma origem, parando um passo antes de
    // qualquer cruzamento de origem. Importante: o frameset MAIS externo do Projudi roda
    // em "projudi.tjpr.jus.br" (sem o "2"), enquanto todo o conteúdo — inclusive o widget
    // com os links de Paralisados/Remessas — fica em frames "projudi2.tjpr.jus.br". Como
    // são origens diferentes, window.top é cross-origin a partir de qualquer frame do
    // Projudi e SEMPRE lança erro ao acessar window.top.document — por isso não dá pra
    // simplesmente começar a busca em window.top (ver todosDocumentosAcessiveis).
    function maiorAncestralAcessivel() {
        let w = window;
        for (let i = 0; i < 10; i++) {
            let pai;
            try { pai = w.parent; } catch (e) { break; }
            if (!pai || pai === w) break;
            try { void pai.document; } catch (e) { break; } // cross-origin — para aqui
            w = pai;
        }
        return w;
    }

    // Procura um link de menu (por URL e/ou texto) no documento atual e em toda a árvore
    // de frames a partir do maior ancestral de mesma origem acessível (ver comentário
    // acima — não dá pra usar window.top direto). Alguns links de menu ficam em widgets
    // aninhados vários níveis abaixo (ex.: o card de "Processos Paralisados" da página
    // inicial), fora do alcance de document/parent isolados — por isso a busca recursiva.
    function todosDocumentosAcessiveis() {
        const vistos = new Set();
        const docs = [];
        function visitar(win) {
            let doc;
            try { doc = win.document; } catch (e) { return; }
            if (!doc || vistos.has(doc)) return;
            vistos.add(doc);
            docs.push(doc);
            let frames;
            try { frames = win.frames; } catch (e) { return; }
            if (!frames) return;
            for (let i = 0; i < frames.length; i++) {
                try { visitar(frames[i]); } catch (e) {}
            }
        }
        visitar(maiorAncestralAcessivel());
        if (!docs.length) docs.push(document);
        return docs;
    }

    function acharLinkMenu(urlRe, textoRe) {
        const docs = todosDocumentosAcessiveis();
        for (const d of docs) {
            for (const a of d.querySelectorAll('a[href]')) {
                if (urlRe && !urlRe.test(a.href)) continue;
                if (textoRe && !textoRe.test((a.textContent || '').trim())) continue;
                return a;
            }
        }
        return null;
    }

    // Variante de acharLinkMenu para o caso de Conclusões: o link fica na mesma
    // conclusao.do do Retorno de Processos Conclusos (não há um texto fixo conhecido
    // para "Conclusões" isoladamente), então em vez de exigir um rótulo específico,
    // pega o primeiro link com URL compatível cujo texto NÃO menciona "retorno" — a
    // mesma distinção já usada para separar os dois relatórios pelo cabeçalho da tabela.
    function acharLinkMenuExcluindo(urlRe, excluirTextoRe) {
        const docs = todosDocumentosAcessiveis();
        for (const d of docs) {
            for (const a of d.querySelectorAll('a[href]')) {
                if (urlRe && !urlRe.test(a.href)) continue;
                const texto = (a.textContent || '').trim();
                if (!texto) continue;
                if (excluirTextoRe && excluirTextoRe.test(texto)) continue;
                return a;
            }
        }
        return null;
    }

    // Alguns cards da página inicial têm vários links com a MESMA URL na mesma célula,
    // cada um precedido de um rótulo em texto puro (não um <a> nem um elemento à parte) —
    // ex.: "Secretaria: <a>74</a> Em Remessa: <a>1221</a> ...". Nesses casos, o texto do
    // próprio link (o número) e a URL (idêntica em todos) não bastam para diferenciar; é
    // preciso ler o texto dos nós irmãos imediatamente anteriores ao link, até topar com
    // outro elemento (o que delimita o rótulo específico daquele link).
    function acharLinkPorRotulo(urlRe, rotuloRe) {
        const docs = todosDocumentosAcessiveis();
        for (const d of docs) {
            for (const a of d.querySelectorAll('a[href]')) {
                if (urlRe && !urlRe.test(a.href)) continue;
                // Ignora os ícones de ajuda ("i" ao lado dos rádios do próprio formulário
                // de filtros) — têm a MESMA URL base (processoBuscaParalisado.do) e o texto
                // do rótulo do rádio como "irmão anterior" também bate no rotuloRe, então
                // sem essa exclusão o script clica no botão de ajuda em vez do link do
                // card da página inicial. Esses ícones sempre têm class "ajaxCalloutHelp…"
                // e terminam a URL em "#" (âncora vazia, sem token de navegação real).
                if (/ajaxCalloutHelp/i.test(a.className) || /#$/.test(a.href)) continue;
                let rotulo = '';
                let node = a.previousSibling;
                while (node && node.nodeType !== 1) {
                    rotulo = (node.textContent || '') + rotulo;
                    node = node.previousSibling;
                }
                if (rotuloRe.test(rotulo)) return a;
            }
        }
        return null;
    }

    // Acha um link que fica na mesma linha (<tr>) de um <label> específico — usado para
    // o número de "Processos Suspensos por Tempo Indeterminado" na página inicial, que
    // (ao contrário do card de Paralisados) tem o rótulo numa célula própria antes do
    // link, sem texto solto entre os dois.
    function acharLinkAoLadoDoLabel(labelRe) {
        const docs = todosDocumentosAcessiveis();
        for (const d of docs) {
            for (const label of d.querySelectorAll('td.label label')) {
                if (!labelRe.test(label.textContent)) continue;
                const tr = label.closest('tr');
                const a = tr && tr.querySelector('a[href]');
                if (a) return a;
            }
        }
        return null;
    }

    function navegarMenu(alvo) {
        let link = null;
        if (alvo === 'juntadas') link = acharLinkMenu(/analisarJuntada\.do/i, null);
        else if (alvo === 'conclusoes') link = acharLinkMenuExcluindo(/conclusao\.do/i, /retorno/i);
        else if (alvo === 'retorno') link = acharLinkMenu(/conclusao\.do/i, /retorno de processos conclusos/i);
        else if (alvo === 'tempomedio') link = acharLinkMenu(/conclusao\/estatistica\.do/i, null);
        else if (alvo === 'paralisados') {
            link = acharLinkPorRotulo(/processoBuscaParalisado\.do/i, /secretaria/i) || acharLinkMenu(/processoBuscaParalisado\.do/i, null);
        }
        else if (alvo === 'remessas') link = acharLinkPorRotulo(/processoBuscaParalisado\.do/i, /em\s+remessa.*exceto\s+processos\s+conclusos/i);
        else if (alvo === 'suspensos') link = acharLinkAoLadoDoLabel(/suspensos\s+por\s+tempo\s+indeterminado/i);
        else if (alvo === 'audiencias') link = acharLinkMenu(/audiencia\/busca\.do/i, /^listagem$/i);
        else if (alvo === 'audienciasdesignadas') link = acharLinkMenu(/audiencia\/pautaAudiencia\.do/i, /^ver\s+pauta\s+de\s+hor[áa]rios$/i);
        else if (alvo === 'audienciasrealizadas') link = acharLinkMenu(/audiencia\/estatistica\.do/i, null);
        // "Apreensões em..." fica no menu "Mesa do Escrivão" (aba #tabItemprefix6) —
        // basta casar pela URL de destino (o rótulo completo varia com a competência).
        else if (alvo === 'apreensoes') link = acharLinkMenu(/processo\/criminal\/apreensao\.do/i, /apreens/i) || acharLinkMenu(/processo\/criminal\/apreensao\.do/i, null);
        else if (alvo === 'inicio') link = acharLinkMenu(null, /^in[íi]cio$/i);
        if (!link) { console.warn('[Auto Projudi] link de menu não encontrado:', alvo); return false; }
        console.log(`[Auto Projudi] navegarMenu("${alvo}") — link encontrado, clicando`);
        link.click();
        return true;
    }

    // Chamado ao concluir a coleta de um relatório (pelo coletor). Marca o próximo estado
    // (próximo relatório da fila, ou "ir_fim") e tenta avançar (o próprio frame do relatório
    // costuma ter o menu; senão o poll do painel assume).
    function avancarAutomacao(cfg) {
        const estado = store.getItem(AUTO_ESTADO);
        const rel = relatorioPorCfg(cfg);
        if (!rel || estado !== 'coletando_' + rel.key) {
            console.log(`[Auto Projudi] avancarAutomacao ignorado — estado="${estado}" cfg=${cfg ? cfg.prefixo : 'null'} rel=${rel ? rel.key : 'null'}`);
            return;
        }
        const fila = lerFilaAutomacao();
        const idx = fila.indexOf(rel.key);
        const prox = idx >= 0 ? fila[idx + 1] : undefined;
        console.log(`[Auto Projudi] avancarAutomacao — "${rel.key}" concluído, próximo="${prox || '(fim)'}"`);
        store.setItem(AUTO_ESTADO, prox ? ('ir_' + prox) : 'ir_fim');
        setTimeout(passoAutomacao, 900);
    }

    // "Pular a extração atual" (botão no painel, visível só com a automação em curso) —
    // usado quando a coleta trava e o usuário não quer esperar/reiniciar tudo. Marca o
    // relatório em andamento como interrompido por erro (ver foiInterrompidoPorErro — o
    // Relatório PDF avisa disso, ao invés de fingir que os dados estão completos) e avança
    // a fila normalmente, como se o relatório tivesse terminado.
    function pularRelatorioAtual() {
        const estado = store.getItem(AUTO_ESTADO);
        if (!estado) return;
        let key = null;
        if (estado.startsWith('coletando_')) key = estado.slice('coletando_'.length);
        else if (estado.startsWith('preenchendo_')) key = estado.slice('preenchendo_'.length);
        else if (estado.startsWith('travado_')) key = estado.slice('travado_'.length);
        if (!key) return;
        const rel = relatorioPorChave(key);
        if (!rel) return;
        if (!confirm(`Pular a extração de "${rel.rotulo}"? Ele vai constar no Relatório PDF como interrompido por erro.`)) return;

        console.warn(`[Auto Projudi] usuário pulou a extração de "${rel.rotulo}" (estado="${estado}")`);
        store.setItem(rel.cfg.prefixo + 'erro', '1');
        store.setItem(rel.cfg.prefixo + 'coletado', '1');
        store.removeItem(rel.cfg.prefixo + 'rodando');
        // Limpa o estado transitório de relatórios com fila própria — senão a próxima
        // tentativa desse relatório retomaria do meio (mês/usuário errado) em vez de
        // recomeçar do zero.
        if (rel.key === 'tempomedio') store.removeItem(CHAVE_FILA_MESES_TM);
        if (rel.key === 'audienciasdesignadas') store.removeItem(CHAVE_PROGRESSO_AD);
        if (rel.key === 'audienciasrealizadas') limparEstadoTransitorioAR();

        const fila = lerFilaAutomacao();
        const idx = fila.indexOf(key);
        const prox = idx >= 0 ? fila[idx + 1] : undefined;
        store.setItem(AUTO_ESTADO, prox ? ('ir_' + prox) : 'ir_fim');
        store.setItem('projudi_auto_lock', String(Date.now()));
        atualizarPainel();
        setTimeout(passoAutomacao, 300);
    }

    // Executa o passo de navegação do estado atual. Pode ser chamado por qualquer frame;
    // uma trava de tempo evita cliques duplicados quando mais de um frame o executa.
    function passoAutomacao() {
        const estado = store.getItem(AUTO_ESTADO);
        if (!estado || estado === 'concluido') return;
        // Durante a coleta ou o preenchimento do formulário, quem conduz é a própria
        // página atual (injetarBotoes / preencherEPesquisarTempoMedio /
        // preencherEPesquisarParalisado). "travado_X" = desistiu de navegar depois de
        // várias tentativas (ver ir_ abaixo) — só sai desse estado com "Limpar" ou
        // retomando a automação do zero.
        if (estado.startsWith('coletando_') || estado.startsWith('preenchendo_') || estado.startsWith('travado_')) return;

        const agora = Date.now();
        const lock = parseInt(store.getItem('projudi_auto_lock') || '0', 10);
        if (agora - lock < 4000) return; // já houve tentativa recente em algum frame

        if (estado === 'ir_fim') {
            store.setItem('projudi_auto_lock', String(agora));
            store.setItem(AUTO_ESTADO, 'concluido');
            store.setItem(CHAVE_AUTO_FIM, String(agora));
            navegarMenu('inicio');
            return;
        }
        if (estado.startsWith('ir_')) {
            const key = estado.slice(3);
            const rel = relatorioPorChave(key);
            if (!rel) { console.warn('[Auto Projudi] relatório desconhecido no estado', estado); return; }
            store.setItem('projudi_auto_lock', String(agora));
            if (navegarMenu(rel.navAlvo)) {
                store.removeItem('projudi_auto_nav_falhas');
                store.setItem(AUTO_ESTADO, rel.precisaPreencher ? ('preenchendo_' + key) : ('coletando_' + key));
                return;
            }
            // Conta tentativas malsucedidas para este alvo — se o link nunca aparecer
            // (ex.: o widget da página inicial não carregou o card esperado), depois de
            // várias tentativas paramos em vez de ficar tentando pra sempre em silêncio
            // (o usuário via isso como o script "travado", sem nenhuma indicação do que
            // fazer). O contador é zerado sempre que o estado muda de alvo.
            const chaveFalhas = 'projudi_auto_nav_falhas';
            const registro = JSON.parse(store.getItem(chaveFalhas) || '{}');
            if (registro.key !== key) { registro.key = key; registro.n = 0; }
            registro.n = (registro.n || 0) + 1;
            store.setItem(chaveFalhas, JSON.stringify(registro));
            console.warn(`[Auto Projudi] tentativa ${registro.n} sem sucesso para "${rel.navAlvo}" — URL atual: ${location.href}`);

            const LIMITE_TENTATIVAS = 8; // ~8 tentativas * 4s de trava = ~32s
            if (registro.n >= LIMITE_TENTATIVAS) {
                store.removeItem(chaveFalhas);
                store.setItem(AUTO_ESTADO, 'travado_' + key);
                console.error(`[Auto Projudi] desistindo de navegar para "${rel.navAlvo}" após ${LIMITE_TENTATIVAS} tentativas — automação parada. Navegue manualmente até "${rel.rotulo}" pela página inicial, ou clique em Limpar para recomeçar.`);
                return;
            }
            if (rel.navAlvo === 'paralisados' || rel.navAlvo === 'remessas' || rel.navAlvo === 'suspensos') {
                // O link de Paralisados/Remessas/Suspensos só existe no card da página
                // inicial — se a automação estiver saindo de outro relatório (ex.:
                // resultados de Paralisados), esse link não está na página atual e
                // navegarMenu falha silenciosamente sem nunca sair dali. Volta à início
                // primeiro; o estado continua "ir_X" para tentar de novo assim que a
                // home carregar.
                console.log(`[Auto Projudi] link de "${rel.navAlvo}" não está nesta página — voltando à início para tentar de lá`);
                navegarMenu('inicio');
            }
        }
    }

    // fila: array de chaves ('juntadas'|'retorno'|'tempomedio') na ordem a executar.
    // periodoTM: id do período pré-configurado para a data inicial do Tempo Médio.
    function iniciarAutomacao(fila, periodoTM) {
        fila = (fila || []).filter(k => relatorioPorChave(k));
        if (!fila.length) { alert('Selecione ao menos um relatório para automatizar.'); return; }
        store.setItem('projudi_auto_fila', JSON.stringify(fila));
        store.setItem('projudi_auto_periodo_tm', periodoTM || '1m');
        // Monta a fila de meses do Tempo Médio uma única vez aqui, no início — nunca dentro
        // do loop mês a mês (ver criarColetor/continuar), senão ela voltaria sempre ao
        // tamanho cheio a cada mês coletado.
        if (fila.includes('tempomedio')) prepararFilaMesesTempoMedio(periodoTM || '1m');
        const primeiro = relatorioPorChave(fila[0]);
        store.setItem(AUTO_ESTADO, primeiro.precisaPreencher ? ('preenchendo_' + primeiro.key) : ('coletando_' + primeiro.key));
        store.setItem('projudi_auto_lock', String(Date.now()));
        store.setItem(CHAVE_AUTO_INICIO, String(Date.now()));
        store.removeItem(CHAVE_AUTO_FIM);
        atualizarPainel();
        setTimeout(() => navegarMenu(primeiro.navAlvo), 300);
    }

    function limparTudoAutomacao() {
        REPORTS_AUTOMACAO.map(r => r.cfg).forEach(c => {
            const n = parseInt(store.getItem(c.prefixo + 'num_paginas') || '0', 10);
            for (let i = 0; i < n; i++) store.removeItem(c.prefixo + 'pagina_' + i);
            store.removeItem(c.prefixo + 'num_paginas');
            store.removeItem(c.prefixo + 'rodando');
            store.removeItem(c.prefixo + 'ts');
            store.removeItem(c.prefixo + 'coletado');
            store.removeItem(c.prefixo + 'erro');
        });
        store.removeItem(AUTO_ESTADO);
        store.removeItem(CHAVE_AUTO_INICIO);
        store.removeItem(CHAVE_AUTO_FIM);
        store.removeItem('projudi_auto_fila');
        store.removeItem('projudi_auto_periodo_tm');
        store.removeItem('projudi_tempomedio_auto_iniciar');
        store.removeItem(CHAVE_FILA_MESES_TM);
        store.removeItem('projudi_paralisado_auto_iniciar');
        store.removeItem('projudi_auto_nav_falhas');
        store.removeItem('projudi_estatisticas_ativos');
        limparEstadoTransitorioAR();
        atualizarPainel();
    }

    // Seções com cfg.mostrarSeVazio (Suspensos, Audiências Pendentes) entram mesmo com
    // zero registros, desde que a coleta tenha rodado até o fim — "zero pendências" é uma
    // informação relevante para o card, diferente de "nunca foi coletado".
    function foiColetado(cfg) { return store.getItem(cfg.prefixo + 'coletado') === '1'; }
    // Marcado por "Pular a extração atual" no painel (ver pularRelatorioAtual) — o
    // relatório foi interrompido antes de terminar; os dados que existem são parciais.
    function foiInterrompidoPorErro(cfg) { return store.getItem(cfg.prefixo + 'erro') === '1'; }

    function baixarPDFConjunto(somenteResumo) {
        const secoes = REPORTS_AUTOMACAO
            .map(r => ({ dados: lerDadosDe(r.cfg.prefixo), cfg: r.cfg }))
            .filter(s => s.dados.length || (s.cfg.mostrarSeVazio && foiColetado(s.cfg)));
        if (!secoes.length) { alert('Nenhum dado coletado ainda.'); return; }
        try {
            gerarPDFConjunto(secoes, somenteResumo);
            // Após exportar o PDF conjunto, limpa tudo automaticamente — evita que uma
            // coleta antiga fique acumulada/misturada com a próxima automação.
            limparTudoAutomacao();
        }
        catch (err) { alert('Erro ao gerar Relatório PDF: ' + err.message); console.error(err); }
    }

    // Calcula o progresso da fila de automação: quantos relatórios já foram coletados por
    // completo (concluidos) e a fração correspondente (0 a 1, com meio ponto de crédito
    // para o relatório em andamento no momento — só conta como completo quando termina).
    function progressoAutomacao(estado, fila) {
        const totalFila = fila.length;
        if (!totalFila) return { concluidos: 0, frac: 0 };
        if (estado === 'inativo' || !estado) return { concluidos: 0, frac: 0 };
        if (estado === 'concluido' || estado === 'ir_fim') return { concluidos: totalFila, frac: 1 };
        let key = null, emAndamento = false;
        if (estado.startsWith('preenchendo_')) { key = estado.slice('preenchendo_'.length); emAndamento = true; }
        else if (estado.startsWith('coletando_')) { key = estado.slice(10); emAndamento = true; }
        else if (estado.startsWith('ir_')) { key = estado.slice(3); emAndamento = false; }
        else if (estado.startsWith('travado_')) { key = estado.slice(8); emAndamento = false; }
        const idx = key ? fila.indexOf(key) : -1;
        if (idx < 0) return { concluidos: 0, frac: 0 };
        const concluidos = idx; // relatórios antes deste na fila já foram coletados
        return { concluidos, frac: (idx + (emAndamento ? 0.5 : 0)) / totalFila };
    }

    // Painel flutuante (na página que tem o menu principal / página inicial)
    function atualizarPainel() {
        const painel = document.getElementById('painel-automacao');
        if (!painel) return;
        const estado = store.getItem(AUTO_ESTADO) || 'inativo';
        const contagens = REPORTS_AUTOMACAO.map(r => ({ r, n: lerDadosDe(r.cfg.prefixo).length }));
        const total = contagens.reduce((s, c) => s + c.n, 0);
        const emCurso = estado !== 'inativo' && estado !== 'concluido';
        const travado = estado.startsWith('travado_');
        const estadoTexto = (() => {
            if (estado === 'inativo') return 'Pronto para iniciar';
            if (estado === 'concluido') return 'Coleta concluída';
            if (estado === 'ir_fim') return 'Finalizando…';
            if (estado.startsWith('preenchendo_')) { const rel = relatorioPorChave(estado.slice('preenchendo_'.length)); return `Preenchendo filtros de <strong>${rel ? rel.rotulo : estado}</strong>…`; }
            if (estado.startsWith('coletando_')) {
                const chave = estado.slice(10);
                const rel = relatorioPorChave(chave);
                let txt = `Coletando <strong>${rel ? rel.rotulo : estado}</strong>…`;
                // Audiências Designadas expande processo a processo (em lotes) e pode
                // demorar bastante em pautas grandes — mostra o progresso em vez de deixar
                // o texto parado sem indicação de que algo está acontecendo.
                if (chave === 'audienciasdesignadas') {
                    const prog = lerProgressoExpansaoAD();
                    if (prog && prog.total > 0) txt += ` (${prog.processados}/${prog.total} linha(s) da pauta)`;
                }
                // Audiências Realizadas pesquisa usuário por usuário (uma pesquisa =
                // um reload da página) — mesma ideia, mostra o progresso da fila.
                if (chave === 'audienciasrealizadas') {
                    const prog = lerProgressoAR();
                    if (prog && prog.total > 0) txt += ` (${prog.processados}/${prog.total} usuário(s))`;
                }
                return txt;
            }
            if (estado.startsWith('ir_')) { const rel = relatorioPorChave(estado.slice(3)); return `Indo para <strong>${rel ? rel.rotulo : estado}</strong>…`; }
            if (travado) {
                const rel = relatorioPorChave(estado.slice(8));
                const nome = rel ? rel.rotulo : estado.slice(8);
                return `Não encontrou o link para "<strong>${nome}</strong>" — navegue manualmente até lá pela página inicial, ou clique em Limpar para recomeçar`;
            }
            return estado;
        })();

        const dot = painel.querySelector('.pa-dot');
        painel.querySelector('.pa-state-txt').innerHTML = estadoTexto;
        if (dot) dot.classList.toggle('on', emCurso);
        if (dot) dot.classList.toggle('alerta', travado);
        painel.querySelectorAll('.pa-count').forEach(el => {
            const c = contagens.find(c => c.r.key === el.dataset.key);
            const n = el.querySelector('.n');
            if (c && n) n.textContent = String(c.n);
        });
        painel.querySelector('#pa-iniciar').disabled = emCurso;
        painel.querySelector('#pa-pdf').disabled = total === 0;
        // "Pular extração atual" só aparece com a automação em curso — é a válvula de
        // escape para quando a coleta trava (ver pularRelatorioAtual).
        const btnPular = painel.querySelector('#pa-pular');
        if (btnPular) btnPular.style.display = emCurso ? '' : 'none';

        // Tempo decorrido: enquanto em curso, atualiza a cada 2s (mesmo intervalo que
        // chama atualizarPainel); depois de concluído, mostra o tempo total fixo da
        // última rodada.
        const elTempo = painel.querySelector('.pa-tempo');
        if (elTempo) {
            const inicio = parseInt(store.getItem(CHAVE_AUTO_INICIO) || '0', 10);
            const fim = parseInt(store.getItem(CHAVE_AUTO_FIM) || '0', 10);
            if (inicio && emCurso) {
                elTempo.style.display = '';
                elTempo.textContent = `Tempo decorrido: ${formatarDuracao(Date.now() - inicio)}`;
            } else if (inicio && fim && estado === 'concluido') {
                elTempo.style.display = '';
                elTempo.textContent = `Extração concluída em ${formatarDuracao(fim - inicio)}`;
            } else {
                elTempo.style.display = 'none';
            }
        }
        // "Limpar" nunca é desabilitado — é o botão de resgate caso a automação trave
        // num estado intermediário (evita o usuário ficar sem forma de apagar e recomeçar).
        painel.querySelectorAll('.pa-check, #pa-periodo-tm, #pa-marcar-tudo, #pa-desmarcar-tudo').forEach(el => {
            el.disabled = emCurso;
        });

        // Barra de progresso: proporção da fila já coletada (fica escondida quando não há
        // fila registrada — ex.: antes da primeira automação, ou depois de "Limpar").
        const fila = lerFilaAutomacao();
        const { concluidos, frac } = progressoAutomacao(estado, fila);
        const pct = Math.round(frac * 100);
        const wrap = painel.querySelector('.pa-progress');
        const barra = painel.querySelector('.pa-progress-bar');
        const label = painel.querySelector('.pa-progress-lbl');
        if (wrap && barra && label) {
            wrap.style.display = fila.length ? '' : 'none';
            barra.style.width = pct + '%';
            barra.classList.toggle('pa-progress-completo', estado === 'concluido');
            label.innerHTML = `<span>${concluidos} de ${fila.length} concluído(s)</span><span>${pct}%</span>`;
        }
    }

    // Categorias do painel: Cível-Geral é a base (todos os relatórios já existentes);
    // as demais herdam os mesmos itens e ganham uma seção própria para relatórios
    // específicos — ainda vazia (placeholder) até serem definidos. Só Cível-Geral e
    // Crime por enquanto; Família/Infância entra depois, no mesmo padrão.
    const CATEGORIAS_PAINEL = [
        { id: 'civel', rotulo: 'Cível-Geral' },
        { id: 'crime', rotulo: 'Crime' },
    ];
    const CHAVE_CATEGORIA_PAINEL = 'projudi_painel_categoria';

    function injetarPainel() {
        if (document.getElementById('painel-automacao')) return;
        // Só na página que hospeda o menu principal (evita duplicar em outros frames).
        // Tentativa anterior de inserir como "quarto menu" da página inicial (ao lado de
        // Dados do Juízo/Processos Ativos/Últimas Mensagens) não deu certo na prática —
        // volta a ser um painel flutuante, como antes.
        if (!document.querySelector('#main-menu')) return;
        // Não injeta sobre a tela de resultados de um relatório
        if (detectarConfig()) return;

        const painel = document.createElement('div');
        painel.id = 'painel-automacao';
        // Itens do Cível-Geral (sem categoriaEspecifica) valem para qualquer categoria —
        // ficam sempre visíveis nos grupos Cartório/Gabinete, independente da aba. Itens
        // com categoriaEspecifica só aparecem na seção própria da categoria correspondente
        // (ver aplicarCategoria) — hoje só "Audiências Pendentes" em Crime.
        const itensCivel = REPORTS_AUTOMACAO.filter(r => !r.categoriaEspecifica);
        const itensEspecificos = (catId) => REPORTS_AUTOMACAO.filter(r => r.categoriaEspecifica === catId);
        const itensVisiveis = (catId) => itensCivel.concat(itensEspecificos(catId));

        // Todos os relatórios vêm marcados por padrão, exceto Tempo Médio (precisa ser
        // habilitado explicitamente — ver seletor de período ao lado).
        function linhaChecklistItem(r) {
            const seletorPeriodo = r.key === 'tempomedio'
                ? `<select id="pa-periodo-tm" class="sel-periodo" title="Quantos meses completos buscar (sempre em pesquisas separadas por mês)">${
                    PERIODOS_TEMPOMEDIO.map(p => `<option value="${p.id}"${p.id === '1m' ? ' selected' : ''}>${p.rotulo}</option>`).join('')
                  }</select>`
                : '';
            return `
                    <label class="pa-item">
                        <input type="checkbox" class="pa-check" data-key="${r.key}" ${r.key === 'tempomedio' ? '' : 'checked'}> ${r.rotulo}${seletorPeriodo}
                    </label>`;
        }
        const linhasGrupos = GRUPOS_AUTOMACAO.map(g => {
            const itens = itensCivel.filter(r => r.dominio === g.chave).map(linhaChecklistItem).join('');
            return `
                <div class="pa-group">
                    <p class="pa-group-lbl">${g.rotulo}</p>
                    ${itens}
                </div>`;
        }).join('');
        const montarContagem = (catId) => itensVisiveis(catId).map(r => `
                    <div class="pa-count" data-key="${r.key}"><span class="l">${r.curto}</span><span class="n">0</span></div>`).join('');
        // Enquanto a categoria não tiver nenhum item próprio ainda definido, mostra um
        // espaço reservado em vez de uma seção vazia.
        const montarGrupoEspecifico = (cat) => {
            const itens = itensEspecificos(cat.id);
            return itens.length ? itens.map(linhaChecklistItem).join('')
                : `<div class="pa-placeholder">Itens específicos desta categoria — a definir</div>`;
        };
        const categoriaSalva = store.getItem(CHAVE_CATEGORIA_PAINEL) || 'civel';
        const categoriaInicial = CATEGORIAS_PAINEL.some(c => c.id === categoriaSalva) ? categoriaSalva : 'civel';
        const linhasAbas = CATEGORIAS_PAINEL.map(c => `
                    <button class="pa-tab${c.id === categoriaInicial ? ' active' : ''}" type="button" data-categoria="${c.id}">${c.rotulo}</button>`).join('');
        painel.innerHTML = `
            <div class="pa-head">
                <span class="pa-titulo">Automação de relatórios</span>
                <div class="pa-icons">
                    <button class="pa-icon-btn pa-btn-colapsar" type="button" title="Expandir">▼</button>
                    <button class="pa-icon-btn pa-btn-fechar" type="button" title="Fechar">✕</button>
                </div>
            </div>
            <div class="pa-body" style="display:none;">
                <div class="pa-tabs">${linhasAbas}</div>
                <div class="pa-state-row">
                    <span class="pa-dot"></span>
                    <span class="pa-state-txt">—</span>
                </div>
                <div class="pa-tempo" style="display:none;"></div>
                <div class="pa-progress" style="display:none;">
                    <div class="pa-progress-track"><div class="pa-progress-bar"></div></div>
                    <div class="pa-progress-lbl">—</div>
                </div>
                <div class="pa-counts">${montarContagem(categoriaInicial)}</div>
                ${linhasGrupos}
                <div class="pa-group pa-group-especifico" style="display:none;">
                    <p class="pa-group-lbl especifico"></p>
                    <div class="pa-group-conteudo"></div>
                </div>
                <div class="pa-links">
                    <button id="pa-marcar-tudo" class="pa-link" type="button">Marcar tudo</button>
                    <button id="pa-desmarcar-tudo" class="pa-link" type="button">Desmarcar tudo</button>
                </div>
                <label class="pa-resumo" title="Gera o Relatório PDF só com os resumos (KPIs e gráficos) de cada relatório, sem as tabelas discriminadas">
                    <input type="checkbox" id="pa-somente-resumo"> Só resumo (sem tabelas discriminadas)
                </label>
                <label class="pa-resumo" title="Reexibe os botões de Extrair/Baixar/PDF/Limpar em cada tela de relatório (fora do painel de automação)">
                    <input type="checkbox" id="pa-mostrar-botoes"${mostrarBotoesIndividuais() ? ' checked' : ''}> Mostrar botões individuais nos relatórios
                </label>
                <div class="pa-actions">
                    <button id="pa-iniciar" class="pa-btn pa-btn-primary" type="button" title="Extrai os relatórios marcados automaticamente">▶ Automatizar</button>
                    <div class="pa-btn-row">
                        <button id="pa-pdf" class="pa-btn pa-btn-secondary" type="button" title="Gera um PDF único com os relatórios já coletados">⬇ Relatório PDF</button>
                        <button id="pa-limpar" class="pa-btn pa-btn-ghost" type="button" title="Apaga os dados acumulados de todos os relatórios">Limpar</button>
                    </div>
                    <button id="pa-pular" class="pa-btn pa-btn-ghost pa-btn-alerta" type="button" style="display:none;" title="Pula a extração do relatório atual (use em caso de travamento) — ele consta no Relatório PDF como interrompido por erro">⏭ Pular extração atual</button>
                </div>
                <div class="pa-dica">Rode em cada Atuação para acumular várias competências antes de gerar o Relatório PDF.</div>
            </div>`;
        document.body.appendChild(painel);
        painel.querySelector('#pa-iniciar').onclick = () => {
            // Itens de outra categoria ficam com o checkbox oculto (display:none no
            // .pa-group), não desmarcado — sem esse filtro, marcar um item específico,
            // trocar de aba e clicar Automatizar rodaria um relatório invisível na tela.
            const fila = [...painel.querySelectorAll('.pa-check:checked')]
                .filter(c => { const grupo = c.closest('.pa-group'); return !grupo || grupo.style.display !== 'none'; })
                .map(c => c.dataset.key).filter(Boolean);
            // #pa-periodo-tm só existe quando o Tempo Médio está ativo em REPORTS_AUTOMACAO.
            const periodoSelTM = painel.querySelector('#pa-periodo-tm');
            const periodoTM = periodoSelTM ? periodoSelTM.value : '1m';
            iniciarAutomacao(fila, periodoTM);
        };
        painel.querySelector('#pa-pdf').onclick = () => {
            const chk = painel.querySelector('#pa-somente-resumo');
            baixarPDFConjunto(!!(chk && chk.checked));
        };
        painel.querySelector('#pa-limpar').onclick = limparTudoAutomacao;
        painel.querySelector('#pa-pular').onclick = pularRelatorioAtual;
        painel.querySelector('#pa-mostrar-botoes').onchange = (ev) => {
            store.setItem(CHAVE_MOSTRAR_BOTOES, ev.target.checked ? '1' : '0');
        };
        painel.querySelector('#pa-marcar-tudo').onclick = () => painel.querySelectorAll('.pa-check').forEach(c => { c.checked = true; });
        painel.querySelector('#pa-desmarcar-tudo').onclick = () => painel.querySelectorAll('.pa-check').forEach(c => { c.checked = false; });
        painel.querySelector('.pa-btn-colapsar').onclick = () => {
            const body = painel.querySelector('.pa-body');
            const btn = painel.querySelector('.pa-btn-colapsar');
            const recolhido = body.style.display === 'none';
            body.style.display = recolhido ? '' : 'none';
            btn.textContent = recolhido ? '▲' : '▼';
            btn.title = recolhido ? 'Recolher' : 'Expandir';
        };
        painel.querySelector('.pa-btn-fechar').onclick = () => painel.remove();

        // Troca de categoria: mostra a seção de itens específicos (checkboxes de verdade
        // quando já existem, senão o espaço reservado) e refaz a grade de contagens para
        // incluir só os itens visíveis na aba atual (Cível-Geral + os específicos dela).
        function aplicarCategoria(id) {
            const cat = CATEGORIAS_PAINEL.find(c => c.id === id) || CATEGORIAS_PAINEL[0];
            painel.querySelectorAll('.pa-tab').forEach(btn => {
                btn.classList.toggle('active', btn.dataset.categoria === cat.id);
            });
            const grupoEspecifico = painel.querySelector('.pa-group-especifico');
            const ehCivel = cat.id === 'civel';
            grupoEspecifico.style.display = ehCivel ? 'none' : '';
            grupoEspecifico.querySelector('.pa-group-lbl').textContent = cat.rotulo;
            grupoEspecifico.querySelector('.pa-group-conteudo').innerHTML = montarGrupoEspecifico(cat);
            painel.querySelector('.pa-counts').innerHTML = montarContagem(cat.id);
            store.setItem(CHAVE_CATEGORIA_PAINEL, cat.id);
            atualizarPainel(); // repopula os números da grade de contagem recém-trocada
        }
        painel.querySelectorAll('.pa-tab').forEach(btn => {
            btn.onclick = () => aplicarCategoria(btn.dataset.categoria);
        });
        aplicarCategoria(categoriaInicial);

        atualizarPainel();
    }

    GM_addStyle(`
        #painel-automacao {
            position: fixed; top: 8px; right: 8px; z-index: 999999; width: 308px;
            background: #FFFFFF; border: 1px solid #DEDDD6; border-radius: 8px;
            box-shadow: 0 6px 20px rgba(26,26,26,.18); overflow: hidden;
            font-family: "Public Sans", Verdana, Arial, sans-serif; color: #1A1A1A;
        }
        #painel-automacao .pa-head {
            display: flex; align-items: center; justify-content: space-between;
            padding: 10px 11px; border-bottom: 1px solid #DEDDD6;
            border-left: 3px solid #3A5A7D; background: #F4F4F1;
        }
        #painel-automacao .pa-titulo { font-size: .82em; font-weight: 700; color: #1A1A1A; }
        #painel-automacao .pa-icons { display: flex; gap: 4px; }
        #painel-automacao .pa-icon-btn {
            width: 20px; height: 20px; border: 1px solid #DEDDD6; border-radius: 5px;
            background: #FFFFFF; cursor: pointer; color: #52514E; font-size: .68em;
            line-height: 1; padding: 0;
        }
        #painel-automacao .pa-icon-btn:hover { background: #F4F4F1; }

        #painel-automacao .pa-tabs {
            display: flex; gap: 2px; padding: 8px 11px 0; background: #F4F4F1; border-bottom: 1px solid #DEDDD6;
            /* .pa-tabs mora dentro de .pa-body (para sumir junto ao colapsar — ver
               pa-btn-colapsar) mas mantém a aparência de faixa colada nas bordas do painel,
               cancelando o padding do .pa-body com margem negativa. */
            margin: -11px -11px 8px;
        }
        #painel-automacao .pa-tab {
            flex: 1; border: none; background: none; padding: 7px 4px 8px; font-size: .68em; font-weight: 600;
            color: #82807A; cursor: pointer; border-bottom: 2px solid transparent; text-align: center;
            font-family: inherit;
        }
        #painel-automacao .pa-tab:hover { color: #52514E; }
        #painel-automacao .pa-tab.active { color: #3A5A7D; border-bottom-color: #3A5A7D; }

        #painel-automacao .pa-body { padding: 11px; }

        #painel-automacao .pa-state-row { display: flex; align-items: center; gap: 7px; margin-bottom: 8px; }
        #painel-automacao .pa-dot {
            width: 7px; height: 7px; border-radius: 50%; background: #C3C2B7; flex: none;
            box-shadow: 0 0 0 3px rgba(195,194,183,.2);
        }
        #painel-automacao .pa-dot.on { background: #9C742E; box-shadow: 0 0 0 3px rgba(156,116,46,.16); }
        #painel-automacao .pa-dot.alerta { background: #923A3A; box-shadow: 0 0 0 3px rgba(146,58,58,.16); }
        #painel-automacao .pa-state-txt { font-size: .72em; color: #52514E; line-height: 1.35; }
        #painel-automacao .pa-state-txt strong { color: #1A1A1A; font-weight: 600; }

        #painel-automacao .pa-progress { margin-bottom: 10px; }
        #painel-automacao .pa-progress-track { background: #DEDDD6; border-radius: 4px; height: 6px; overflow: hidden; }
        #painel-automacao .pa-progress-bar {
            background: #3A5A7D; height: 100%; width: 0%; border-radius: 4px; transition: width .4s ease;
        }
        #painel-automacao .pa-progress-bar.pa-progress-completo { background: #527467; }
        #painel-automacao .pa-progress-lbl { display: flex; justify-content: space-between; font-size: .66em; color: #82807A; margin-top: 4px; }
        #painel-automacao .pa-tempo { font-size: .66em; color: #82807A; margin: -4px 0 8px; }

        #painel-automacao .pa-counts {
            display: grid; grid-template-columns: repeat(2, minmax(0,1fr)); gap: 4px 10px;
            margin-bottom: 12px; padding: 8px 9px; background: #F4F4F1; border: 1px solid #DEDDD6; border-radius: 6px;
        }
        #painel-automacao .pa-count { display: flex; justify-content: space-between; font-size: .7em; }
        #painel-automacao .pa-count .n { font-weight: 700; color: #1A1A1A; }
        #painel-automacao .pa-count .l { color: #52514E; }

        #painel-automacao .pa-group { margin-bottom: 10px; }
        #painel-automacao .pa-group-lbl {
            display: flex; align-items: center; gap: 6px; font-size: .66em; font-weight: 700;
            letter-spacing: .05em; text-transform: uppercase; color: #82807A; margin: 0 0 5px;
        }
        #painel-automacao .pa-group-lbl::after { content: ""; flex: 1; height: 1px; background: #DEDDD6; }
        #painel-automacao .pa-group-lbl.especifico { color: #3A5A7D; }
        #painel-automacao .pa-item { font-size: .76em; color: #1A1A1A; display: flex; align-items: center; gap: 6px; padding: 2px 0; }
        #painel-automacao .pa-item input[type="checkbox"] { margin: 0; }
        #painel-automacao .pa-item .sel-periodo, #painel-automacao .pa-item .projudi-select { margin-left: auto; padding: 1px 4px; font-size: .92em; }
        #painel-automacao .pa-placeholder {
            display: flex; align-items: center; gap: 6px; padding: 8px 9px; background: #FAFAF8;
            border: 1px dashed #C3C2B7; border-radius: 6px; font-size: .7em; color: #82807A; font-style: italic;
        }

        #painel-automacao .pa-links { display: flex; gap: 10px; margin-bottom: 8px; }
        #painel-automacao .pa-link {
            background: none; border: none; padding: 0; cursor: pointer;
            font-size: .68em; font-weight: 600; color: #3A5A7D;
        }
        #painel-automacao .pa-link:disabled { color: #82807A; cursor: not-allowed; }

        #painel-automacao .pa-resumo {
            display: flex; align-items: center; gap: 7px; font-size: .72em; color: #52514E;
            padding: 7px 9px; background: #F4F4F1; border: 1px solid #DEDDD6; border-radius: 6px; margin-bottom: 10px;
        }

        #painel-automacao .pa-actions { display: flex; flex-direction: column; gap: 6px; margin-bottom: 8px; }
        #painel-automacao .pa-btn {
            border-radius: 6px; padding: 7px 10px; font-size: .78em; font-weight: 600;
            cursor: pointer; border: 1px solid transparent; flex: 1;
        }
        #painel-automacao .pa-btn:disabled { opacity: .5; cursor: not-allowed; }
        #painel-automacao .pa-btn-primary { background: #3A5A7D; border-color: #2E4A69; color: #fff; }
        #painel-automacao .pa-btn-secondary { background: #FFFFFF; color: #3A5A7D; border-color: #3A5A7D; }
        #painel-automacao .pa-btn-ghost { background: none; color: #923A3A; border-color: #DEDDD6; font-weight: 500; }
        #painel-automacao .pa-btn-alerta { width: 100%; border-color: #E3B98A; background: #FBF2E7; color: #9C742E; }
        #painel-automacao .pa-btn-row { display: flex; gap: 6px; }

        #painel-automacao .pa-dica { font-size: .64em; color: #82807A; line-height: 1.4; border-top: 1px solid #DEDDD6; padding-top: 8px; }
    `);

    // Cada chamada isolada num try/catch: um erro (ex.: numa página com estrutura
    // inesperada) não pode derrubar as chamadas seguintes — em especial passoAutomacao(),
    // sem a qual a automação fica parada sem nenhum aviso ao usuário.
    function chamarSeguro(fn, nome) {
        try { fn(); }
        catch (err) { console.error(`[Projudi] erro em ${nome}:`, err); }
    }

    function bootstrap() {
        console.log(`[Projudi] bootstrap — URL: ${location.href}`);
        chamarSeguro(gravarProcessosAtivosSeDisponivel, 'gravarProcessosAtivosSeDisponivel'); // nº de processos ativos, só existe na página inicial
        chamarSeguro(injetarBotoes, 'injetarBotoes');   // botões nos relatórios (buttonBar)
        chamarSeguro(injetarPainel, 'injetarPainel');   // painel de automação (só na página inicial)
        chamarSeguro(passoAutomacao, 'passoAutomacao'); // avança a automação se estiver em estado de navegação
        // Conduz a navegação da automação em QUALQUER página (não só na página inicial,
        // onde fica o painel) — a coleta ocorre no frame de conteúdo, que pode não ser o
        // mesmo frame com o link do próximo relatório; sem esse poll rodando em toda
        // página, uma falha na tentativa imediata (ex.: frame ainda carregando) deixava a
        // automação parada até uma ação manual do usuário.
        setInterval(() => { chamarSeguro(atualizarPainel, 'atualizarPainel'); chamarSeguro(passoAutomacao, 'passoAutomacao'); }, 2000);
    }

    if (document.readyState === 'loading') {
        document.addEventListener('DOMContentLoaded', bootstrap);
    } else {
        bootstrap();
    }
})();
