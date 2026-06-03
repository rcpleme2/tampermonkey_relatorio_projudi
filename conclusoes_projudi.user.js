// ==UserScript==
// @name         Exportar Conclusões Projudi para Excel
// @namespace    https://projudi2.tjpr.jus.br/
// @version      4.0
// @description  Coleta as conclusões de várias páginas e de várias Atuações, acumula e exporta tudo numa planilha Excel
// @author       rcpleme2
// @match        https://projudi2.tjpr.jus.br/projudi/processo/conclusao.do*
// @require      https://cdn.sheetjs.com/xlsx-latest/package/dist/xlsx.full.min.js
// @grant        GM_addStyle
// @grant        GM_download
// ==/UserScript==

(function () {
    'use strict';

    // Usamos localStorage (não sessionStorage) para que os dados persistam mesmo ao
    // navegar para a tela de troca de Área de Atuação ou ao fechar/reabrir a aba.
    // Cada página coletada é gravada em sua própria chave (KEY_PAGINA_PREFIXO + índice),
    // evitando reserializar o blob inteiro a cada reload. O índice cresce continuamente
    // entre páginas E entre atuações — assim a coleta sempre ACUMULA, e cada registro já
    // carrega sua própria Atuação, permitindo juntar tudo no final.
    const store = window.localStorage;

    const KEY_RODANDO        = 'projudi_export_rodando';  // "1" = coletando páginas da atuação atual
    const KEY_NUM_PAGINAS    = 'projudi_export_num_paginas';
    const KEY_PAGINA_PREFIXO = 'projudi_export_pagina_';
    const KEY_TIMESTAMP      = 'projudi_export_ts';       // marca de tempo da última atividade

    // Coleta "em andamento" sem atividade por mais que isto é considerada obsoleta
    // (ex.: execução interrompida) e a flag de execução é descartada — sem apagar os dados.
    const STALE_MS = 2 * 60 * 1000; // 2 minutos

    function marcarAtividade() {
        store.setItem(KEY_TIMESTAMP, String(Date.now()));
    }

    function coletaObsoleta() {
        const ts = parseInt(store.getItem(KEY_TIMESTAMP) || '0', 10);
        return !ts || (Date.now() - ts) > STALE_MS;
    }

    // ── Persistência dos dados coletados ────────────────────────────────────────

    function limparTudoArmazenado() {
        const n = parseInt(store.getItem(KEY_NUM_PAGINAS) || '0', 10);
        for (let i = 0; i < n; i++) store.removeItem(KEY_PAGINA_PREFIXO + i);
        store.removeItem(KEY_NUM_PAGINAS);
        store.removeItem(KEY_RODANDO);
        store.removeItem(KEY_TIMESTAMP);
    }

    // Acrescenta os dados de uma página. Retorna o total de registros acumulados.
    function adicionarPagina(dadosPagina) {
        const idx = parseInt(store.getItem(KEY_NUM_PAGINAS) || '0', 10);
        store.setItem(KEY_PAGINA_PREFIXO + idx, JSON.stringify(dadosPagina));
        store.setItem(KEY_NUM_PAGINAS, String(idx + 1));
        return contarRegistros();
    }

    function contarRegistros() {
        return lerTodosOsDados().length;
    }

    function contarAtuacoes() {
        const set = new Set(lerTodosOsDados().map(d => d.atuacao || ''));
        set.delete('');
        return set.size;
    }

    // Desembrulha um valor lido do storage até obter um array. Aceita JSON serializado
    // uma ou mais vezes, conforme o ambiente do Tampermonkey/navegador serializa o valor.
    function normalizarParaArray(valor) {
        let v = valor;
        let tentativas = 0;
        while (typeof v === 'string' && tentativas < 5) {
            try { v = JSON.parse(v); } catch (e) { break; }
            tentativas++;
        }
        return Array.isArray(v) ? v : null;
    }

    // Lê e concatena os dados de todas as páginas (de todas as atuações coletadas).
    function lerTodosOsDados() {
        const n = parseInt(store.getItem(KEY_NUM_PAGINAS) || '0', 10);
        let dados = [];
        for (let i = 0; i < n; i++) {
            const bruto = store.getItem(KEY_PAGINA_PREFIXO + i);
            if (!bruto) continue;
            const parte = normalizarParaArray(bruto);
            if (parte) dados = dados.concat(parte);
            else console.error('[Exportar Projudi] página não pôde ser lida no índice', i);
        }
        return dados;
    }

    GM_addStyle(`
        .projudi-btn {
            color: white;
            border: 1px solid #145214;
            padding: 3px 10px;
            cursor: pointer;
            font-size: 0.85em;
            font-family: Verdana, Arial, sans-serif;
            font-weight: bold;
            margin-left: 6px;
            border-radius: 3px;
        }
        #btn-coletar { background-color: #1e6b1e; border-color: #145214; }
        #btn-baixar  { background-color: #b8860b; border-color: #8a6508; }
        #btn-limpar  { background-color: #8a3b3b; border-color: #6e2f2f; }
        .projudi-btn:disabled { background-color: #999; border-color: #777; cursor: not-allowed; }
        #exportar-status {
            font-size: 0.8em;
            color: #555;
            margin-left: 8px;
            font-family: Verdana, Arial, sans-serif;
        }
    `);

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

    // ── Coleta os dados da tabela visível na página atual ──────────────────────

    function coletarPaginaAtual() {
        const atuacao = lerAtuacao();
        const linhas = document.querySelectorAll('table.resultTable tbody tr');
        const dados = [];
        linhas.forEach(tr => {
            const tds = tr.querySelectorAll('td');
            if (tds.length < 8) return;

            const dtRemessa   = (tds[1].textContent || '').trim();

            const emProcesso  = tds[2].querySelector('em');
            const divClasse   = tds[2].querySelector('div');
            const processo    = emProcesso ? emProcesso.textContent.trim() : tds[2].textContent.trim();
            const classe      = divClasse  ? divClasse.textContent.trim()  : '';

            const seq         = (tds[3].textContent || '').trim();
            const tipoConclusao = (tds[4].textContent || '').trim();
            const privativa   = (tds[5].textContent || '').trim();
            const responsavel = (tds[6].textContent || '').trim();
            const preAnalise  = (tds[7].textContent || '').trim();
            const agrupador   = tds[8] ? (tds[8].textContent || '').trim() : '';

            dados.push({ atuacao, dtRemessa, processo, classe, seq, tipoConclusao, privativa, responsavel, preAnalise, agrupador });
        });
        return dados;
    }

    function temProximaPagina() {
        return !!document.querySelector('a.arrowNextOn');
    }

    function numeroPaginaAtual() {
        const bold = document.querySelector('div#navigator b');
        return bold ? parseInt(bold.textContent.trim(), 10) : 1;
    }

    function totalRegistrosPagina() {
        const navLeft = document.querySelector('div#navigator div.navLeft');
        if (!navLeft) return null;
        const match = navLeft.textContent.match(/(\d+)\s+registro/);
        return match ? parseInt(match[1], 10) : null;
    }

    // ── Geração e download do arquivo Excel ─────────────────────────────────────

    function gerarEbaixarExcel(dados) {
        const cabecalhos = ['Atuação', 'Dt. Remessa', 'Processo', 'Classe', 'Seq.', 'Tipo de Conclusão', 'Privativa', 'Responsável', 'Pré-análise', 'Agrupador'];
        const linhas = dados.map(d => [
            d.atuacao, d.dtRemessa, d.processo, d.classe, d.seq,
            d.tipoConclusao, d.privativa, d.responsavel, d.preAnalise, d.agrupador,
        ]);

        const wb = XLSX.utils.book_new();
        const ws = XLSX.utils.aoa_to_sheet([cabecalhos, ...linhas]);
        ws['!cols'] = [
            { wch: 24 }, { wch: 18 }, { wch: 26 }, { wch: 18 }, { wch: 6 },
            { wch: 30 }, { wch: 10 }, { wch: 30 }, { wch: 14 }, { wch: 20 },
        ];
        XLSX.utils.book_append_sheet(wb, ws, 'Conclusões');

        const nomeArquivo = `conclusoes_projudi_${new Date().toISOString().slice(0, 10)}.xlsx`;
        const wbout = XLSX.write(wb, { bookType: 'xlsx', type: 'array' });
        const blob = new Blob([wbout], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
        const url = URL.createObjectURL(blob);
        const a = document.createElement('a');
        a.href = url;
        a.download = nomeArquivo;
        document.body.appendChild(a);
        a.click();
        setTimeout(() => { URL.revokeObjectURL(url); a.remove(); }, 2000);
    }

    function atualizarStatus(msg) {
        const el = document.getElementById('exportar-status');
        if (el) el.textContent = msg;
    }

    // ── Fluxo de coleta (acumulando entre atuações) ─────────────────────────────
    //
    // 1. "Coletar esta Atuação": percorre todas as páginas da atuação atual,
    //    ACRESCENTANDO aos dados já coletados (não apaga nada).
    // 2. O usuário troca de Atuação no sistema e clica "Coletar" de novo.
    // 3. Quando quiser, clica "Baixar planilha" para juntar tudo e exportar.
    // 4. "Limpar" zera os dados acumulados para começar do zero.

    function iniciarColetaAtuacao() {
        store.setItem(KEY_RODANDO, '1');
        marcarAtividade();
        continuarColeta();
    }

    function continuarColeta() {
        desabilitarBotoes(true);
        marcarAtividade();

        const dadosPagina = coletarPaginaAtual();
        let totalAcumulado;
        try {
            totalAcumulado = adicionarPagina(dadosPagina);
        } catch (err) {
            store.removeItem(KEY_RODANDO);
            atualizarStatus(`Erro ao armazenar dados (${err.name}). Os dados já coletados foram mantidos.`);
            console.error('[Exportar Projudi]', err);
            renderizarUI();
            return;
        }

        const atuacao   = lerAtuacao() || '(sem atuação)';
        const pagina    = numeroPaginaAtual();
        const total     = totalRegistrosPagina();
        const porPagina = dadosPagina.length || 20;
        const totalPags = total ? Math.ceil(total / porPagina) : '?';

        atualizarStatus(`Coletando "${atuacao}" — página ${pagina} de ${totalPags} — ${totalAcumulado} no total acumulado...`);

        if (temProximaPagina()) {
            // Clica na próxima página — recarrega a página; retoma em continuarColeta() no load.
            document.querySelector('a.arrowNextOn').click();
        } else {
            // Terminou esta atuação: volta ao estado ocioso, mantendo os dados acumulados.
            store.removeItem(KEY_RODANDO);
            store.removeItem(KEY_TIMESTAMP);
            renderizarUI();
            atualizarStatus(`Atuação "${atuacao}" coletada. Acumulado: ${totalAcumulado} registros de ${contarAtuacoes()} atuação(ões). Troque de atuação e colete mais, ou baixe a planilha.`);
        }
    }

    function baixarPlanilha() {
        try {
            const dados = lerTodosOsDados();
            if (!dados.length) {
                atualizarStatus('Nenhum registro coletado para exportar.');
                return;
            }
            gerarEbaixarExcel(dados);
            atualizarStatus(`✓ ${dados.length} registros de ${contarAtuacoes()} atuação(ões) exportados. Use "Limpar" para começar uma nova coleta.`);
        } catch (err) {
            atualizarStatus(`Erro ao gerar planilha: ${err.message}`);
            console.error('[Exportar Projudi]', err);
        }
    }

    function limparColeta() {
        limparTudoArmazenado();
        renderizarUI();
        atualizarStatus('Dados acumulados apagados. Pronto para uma nova coleta.');
    }

    // ── Interface ───────────────────────────────────────────────────────────────

    function desabilitarBotoes(desabilitar) {
        ['btn-coletar', 'btn-baixar', 'btn-limpar'].forEach(id => {
            const b = document.getElementById(id);
            if (b) b.disabled = desabilitar;
        });
    }

    function renderizarUI() {
        const total = contarRegistros();
        const btnColetar = document.getElementById('btn-coletar');
        const btnBaixar  = document.getElementById('btn-baixar');
        const btnLimpar  = document.getElementById('btn-limpar');
        if (!btnColetar) return;

        btnColetar.disabled = false;
        btnColetar.textContent = total > 0 ? 'Coletar mais uma Atuação' : 'Coletar esta Atuação';

        btnBaixar.disabled = total === 0;
        btnBaixar.textContent = `⬇ Baixar planilha (${total})`;

        btnLimpar.disabled = total === 0;

        if (total > 0) {
            atualizarStatus(`Acumulado: ${total} registros de ${contarAtuacoes()} atuação(ões).`);
        } else {
            atualizarStatus('');
        }
    }

    function injetarBotoes() {
        const buttonBar = document.querySelector('table.buttonBar td.buttons');
        if (!buttonBar) return;

        const btnColetar = document.createElement('button');
        btnColetar.id = 'btn-coletar';
        btnColetar.type = 'button';
        btnColetar.className = 'projudi-btn';
        btnColetar.title = 'Percorre todas as páginas da Atuação atual e acrescenta aos dados já coletados';
        btnColetar.onclick = iniciarColetaAtuacao;

        const btnBaixar = document.createElement('button');
        btnBaixar.id = 'btn-baixar';
        btnBaixar.type = 'button';
        btnBaixar.className = 'projudi-btn';
        btnBaixar.title = 'Junta tudo o que foi coletado (todas as atuações) e baixa a planilha';
        btnBaixar.onclick = baixarPlanilha;

        const btnLimpar = document.createElement('button');
        btnLimpar.id = 'btn-limpar';
        btnLimpar.type = 'button';
        btnLimpar.className = 'projudi-btn';
        btnLimpar.title = 'Apaga todos os dados acumulados para começar do zero';
        btnLimpar.textContent = 'Limpar';
        btnLimpar.onclick = limparColeta;

        const status = document.createElement('span');
        status.id = 'exportar-status';

        buttonBar.appendChild(btnColetar);
        buttonBar.appendChild(btnBaixar);
        buttonBar.appendChild(btnLimpar);
        buttonBar.appendChild(status);

        if (store.getItem(KEY_RODANDO) === '1' && !coletaObsoleta()) {
            // Coleta em andamento (atividade recente): retoma após o reload da paginação.
            continuarColeta();
        } else {
            // Limpa apenas a flag de execução presa (mantém os dados acumulados).
            store.removeItem(KEY_RODANDO);
            store.removeItem(KEY_TIMESTAMP);
            renderizarUI();
        }
    }

    if (document.readyState === 'loading') {
        document.addEventListener('DOMContentLoaded', injetarBotoes);
    } else {
        injetarBotoes();
    }
})();
