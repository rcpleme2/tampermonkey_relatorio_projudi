// ==UserScript==
// @name         Exportar Conclusões Projudi para Excel
// @namespace    https://projudi2.tjpr.jus.br/
// @version      1.0
// @description  Exporta todos os registros da tabela de conclusões (todas as páginas) para uma planilha Excel
// @author       rcpleme2
// @match        https://projudi2.tjpr.jus.br/projudi/processo/conclusao.do*
// @require      https://cdn.sheetjs.com/xlsx-latest/package/dist/xlsx.full.min.js
// @grant        GM_addStyle
// ==/UserScript==

(function () {
    'use strict';

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
            color: #333;
            margin-left: 8px;
            font-family: Verdana, Arial, sans-serif;
        }
    `);

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
    }

    function setStatus(msg) {
        const el = document.getElementById('exportar-status');
        if (el) el.textContent = msg;
    }

    function coletarPaginaAtual() {
        const linhas = document.querySelectorAll('table.resultTable tbody tr');
        const dados = [];
        linhas.forEach(tr => {
            const tds = tr.querySelectorAll('td');
            if (tds.length < 8) return;

            const dtRemessa = (tds[1].textContent || '').trim();

            const emProcesso = tds[2].querySelector('em');
            const divClasse = tds[2].querySelector('div');
            const processo = emProcesso ? emProcesso.textContent.trim() : (tds[2].textContent || '').trim();
            const classe = divClasse ? divClasse.textContent.trim() : '';

            const seq = (tds[3].textContent || '').trim();
            const tipoConclusao = (tds[4].textContent || '').trim();
            const privativa = (tds[5].textContent || '').trim();
            const responsavel = (tds[6].textContent || '').trim();
            const preAnalise = (tds[7].textContent || '').trim();
            const agrupador = tds[8] ? (tds[8].textContent || '').trim() : '';

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

    function aguardarCarregamento(paginaAnterior) {
        return new Promise((resolve, reject) => {
            const timeout = setTimeout(() => {
                observer.disconnect();
                reject(new Error('Timeout ao aguardar carregamento da página'));
            }, 15000);

            const observer = new MutationObserver(() => {
                const paginaAtual = numeroPaginaAtual();
                if (paginaAtual !== paginaAnterior) {
                    clearTimeout(timeout);
                    observer.disconnect();
                    // Pequena pausa extra para garantir que o tbody foi totalmente renderizado
                    setTimeout(resolve, 300);
                }
            });

            const navigator = document.getElementById('navigator');
            if (navigator) {
                observer.observe(navigator, { childList: true, subtree: true, characterData: true });
            } else {
                // fallback: observar o body
                observer.observe(document.body, { childList: true, subtree: true });
            }
        });
    }

    async function iniciarExportacao() {
        const btn = document.getElementById('btn-exportar-excel');
        btn.disabled = true;

        const total = totalRegistros();
        const totalStr = total ? ` de ~${Math.ceil(total / 20)} páginas` : '';

        try {
            const todosOsDados = [];
            let pagina = 1;

            while (true) {
                setStatus(`Coletando página ${pagina}${totalStr}...`);
                const dados = coletarPaginaAtual();
                todosOsDados.push(...dados);

                if (!temProximaPagina()) break;

                const paginaAnterior = numeroPaginaAtual();
                document.querySelector('a.arrowNextOn').click();
                await aguardarCarregamento(paginaAnterior);
                pagina++;
            }

            setStatus(`Gerando arquivo com ${todosOsDados.length} registros...`);
            gerarExcel(todosOsDados);
            setStatus(`✓ ${todosOsDados.length} registros exportados.`);
        } catch (err) {
            setStatus(`Erro: ${err.message}`);
            console.error('[Exportar Projudi]', err);
        } finally {
            btn.disabled = false;
        }
    }

    function gerarExcel(dados) {
        const cabecalhos = ['Dt. Remessa', 'Processo', 'Classe', 'Seq.', 'Tipo de Conclusão', 'Privativa', 'Responsável', 'Pré-análise', 'Agrupador'];
        const linhas = dados.map(d => [
            d.dtRemessa,
            d.processo,
            d.classe,
            d.seq,
            d.tipoConclusao,
            d.privativa,
            d.responsavel,
            d.preAnalise,
            d.agrupador,
        ]);

        const wb = XLSX.utils.book_new();
        const ws = XLSX.utils.aoa_to_sheet([cabecalhos, ...linhas]);

        // Larguras de coluna sugeridas
        ws['!cols'] = [
            { wch: 18 }, // Dt. Remessa
            { wch: 26 }, // Processo
            { wch: 18 }, // Classe
            { wch: 6  }, // Seq.
            { wch: 30 }, // Tipo de Conclusão
            { wch: 10 }, // Privativa
            { wch: 30 }, // Responsável
            { wch: 14 }, // Pré-análise
            { wch: 20 }, // Agrupador
        ];

        XLSX.utils.book_append_sheet(wb, ws, 'Conclusões');

        const dataHoje = new Date().toISOString().slice(0, 10);
        XLSX.writeFile(wb, `conclusoes_projudi_${dataHoje}.xlsx`);
    }

    // Aguarda a página terminar de carregar antes de injetar o botão
    if (document.readyState === 'loading') {
        document.addEventListener('DOMContentLoaded', injetarBotao);
    } else {
        injetarBotao();
    }
})();
