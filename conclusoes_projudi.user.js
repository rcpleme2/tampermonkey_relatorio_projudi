// ==UserScript==
// @name         Exportar Conclusões Projudi para Excel
// @namespace    https://projudi2.tjpr.jus.br/
// @version      2.0
// @description  Exporta todos os registros da tabela de conclusões (todas as páginas) para uma planilha Excel
// @author       rcpleme2
// @match        https://projudi2.tjpr.jus.br/projudi/processo/conclusao.do*
// @require      https://cdn.sheetjs.com/xlsx-latest/package/dist/xlsx.full.min.js
// @grant        GM_addStyle
// ==/UserScript==

(function () {
    'use strict';

    // Chaves usadas no sessionStorage para persistir dados entre reloads de página
    const KEY_DADOS    = 'projudi_export_dados';
    const KEY_RODANDO  = 'projudi_export_rodando';

    GM_addStyle(`
        #btn-exportar-excel {
            background-color: #1e6b1e;
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
        #btn-exportar-excel:disabled {
            background-color: #888;
            border-color: #666;
            cursor: not-allowed;
        }
        #exportar-status {
            font-size: 0.8em;
            color: #555;
            margin-left: 8px;
            font-family: Verdana, Arial, sans-serif;
        }
    `);

    // ── Coleta os dados da tabela visível na página atual ──────────────────────

    function coletarPaginaAtual() {
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

            dados.push({ dtRemessa, processo, classe, seq, tipoConclusao, privativa, responsavel, preAnalise, agrupador });
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

    function totalRegistros() {
        const navLeft = document.querySelector('div#navigator div.navLeft');
        if (!navLeft) return null;
        const match = navLeft.textContent.match(/(\d+)\s+registro/);
        return match ? parseInt(match[1], 10) : null;
    }

    // ── Geração do arquivo Excel ───────────────────────────────────────────────

    function gerarExcel(dados) {
        const cabecalhos = ['Dt. Remessa', 'Processo', 'Classe', 'Seq.', 'Tipo de Conclusão', 'Privativa', 'Responsável', 'Pré-análise', 'Agrupador'];
        const linhas = dados.map(d => [
            d.dtRemessa, d.processo, d.classe, d.seq,
            d.tipoConclusao, d.privativa, d.responsavel, d.preAnalise, d.agrupador,
        ]);

        const wb = XLSX.utils.book_new();
        const ws = XLSX.utils.aoa_to_sheet([cabecalhos, ...linhas]);
        ws['!cols'] = [
            { wch: 18 }, { wch: 26 }, { wch: 18 }, { wch: 6 },
            { wch: 30 }, { wch: 10 }, { wch: 30 }, { wch: 14 }, { wch: 20 },
        ];
        XLSX.utils.book_append_sheet(wb, ws, 'Conclusões');

        // Gera o binário e dispara o download via Blob para evitar bloqueio do sandbox do Tampermonkey
        const wbout = XLSX.write(wb, { bookType: 'xlsx', type: 'array' });
        const blob = new Blob([wbout], { type: 'application/octet-stream' });
        const url = URL.createObjectURL(blob);
        const a = document.createElement('a');
        a.href = url;
        a.download = `conclusoes_projudi_${new Date().toISOString().slice(0, 10)}.xlsx`;
        document.body.appendChild(a);
        a.click();
        setTimeout(() => { URL.revokeObjectURL(url); a.remove(); }, 1000);
    }

    // ── Lógica principal ───────────────────────────────────────────────────────
    //
    // A paginação do Projudi faz RELOAD COMPLETO da página a cada troca de página.
    // Por isso usamos sessionStorage para acumular dados entre os reloads:
    //   KEY_RODANDO = "1" enquanto a exportação está em andamento
    //   KEY_DADOS   = JSON com o array acumulado de registros
    //
    // Fluxo ao carregar a página:
    //   1. Se KEY_RODANDO == "1": estamos no meio de uma exportação automática
    //      → coleta dados desta página, salva, verifica se há próxima página
    //      → se sim: clica arrowNextOn (causará reload) e aguarda
    //      → se não: gera Excel, limpa sessionStorage
    //   2. Se KEY_RODANDO != "1": aguarda o usuário clicar "Exportar Excel"
    //      → limpa sessionStorage, seta KEY_RODANDO, e inicia o ciclo acima

    function iniciarExportacao() {
        sessionStorage.setItem(KEY_DADOS, JSON.stringify([]));
        sessionStorage.setItem(KEY_RODANDO, '1');
        continuarExportacao();
    }

    function continuarExportacao() {
        const btn = document.getElementById('btn-exportar-excel');
        if (btn) btn.disabled = true;

        const dadosSalvos = JSON.parse(sessionStorage.getItem(KEY_DADOS) || '[]');
        const dadosPagina = coletarPaginaAtual();
        const dadosAcumulados = dadosSalvos.concat(dadosPagina);
        sessionStorage.setItem(KEY_DADOS, JSON.stringify(dadosAcumulados));

        const pagina    = numeroPaginaAtual();
        const total     = totalRegistros();
        const totalPags = total ? Math.ceil(total / dadosPagina.length || 20) : '?';

        atualizarStatus(`Coletando página ${pagina} de ${totalPags} — ${dadosAcumulados.length} registros até agora...`);

        if (temProximaPagina()) {
            // Clica na próxima página — isso vai recarregar a página;
            // o script vai retomar em continuarExportacao() no próximo load.
            document.querySelector('a.arrowNextOn').click();
        } else {
            // Chegamos à última página — gera o arquivo e limpa o estado
            sessionStorage.removeItem(KEY_RODANDO);
            sessionStorage.removeItem(KEY_DADOS);

            atualizarStatus(`Gerando arquivo com ${dadosAcumulados.length} registros...`);
            gerarExcel(dadosAcumulados);
            atualizarStatus(`✓ ${dadosAcumulados.length} registros exportados.`);
            if (btn) btn.disabled = false;
        }
    }

    function atualizarStatus(msg) {
        const el = document.getElementById('exportar-status');
        if (el) el.textContent = msg;
    }

    // ── Injeção do botão na interface ──────────────────────────────────────────

    function injetarBotao() {
        const buttonBar = document.querySelector('table.buttonBar td.buttons');
        if (!buttonBar) return;

        const btn = document.createElement('button');
        btn.id = 'btn-exportar-excel';
        btn.textContent = 'Exportar Excel';
        btn.title = 'Exporta todos os registros de todas as páginas para .xlsx';
        btn.addEventListener('click', iniciarExportacao);

        const status = document.createElement('span');
        status.id = 'exportar-status';

        buttonBar.appendChild(btn);
        buttonBar.appendChild(status);

        // Se havia uma exportação em andamento (retomada após reload de página), continua
        if (sessionStorage.getItem(KEY_RODANDO) === '1') {
            continuarExportacao();
        }
    }

    if (document.readyState === 'loading') {
        document.addEventListener('DOMContentLoaded', injetarBotao);
    } else {
        injetarBotao();
    }
})();
