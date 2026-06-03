// ==UserScript==
// @name         Exportar Conclusões Projudi para Excel
// @namespace    https://projudi2.tjpr.jus.br/
// @version      3.0
// @description  Exporta todos os registros da tabela de conclusões (todas as páginas) para uma planilha Excel
// @author       rcpleme2
// @match        https://projudi2.tjpr.jus.br/projudi/processo/conclusao.do*
// @require      https://cdn.sheetjs.com/xlsx-latest/package/dist/xlsx.full.min.js
// @grant        GM_addStyle
// @grant        GM_download
// ==/UserScript==

(function () {
    'use strict';

    // Chaves usadas no sessionStorage para persistir o estado entre reloads de página
    const KEY_DADOS   = 'projudi_export_dados';
    const KEY_RODANDO = 'projudi_export_rodando'; // "1" = coletando páginas
    const KEY_PRONTO  = 'projudi_export_pronto';  // "1" = coleta terminada, pronto p/ baixar

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
        #btn-exportar-excel.pronto { background-color: #b8860b; border-color: #8a6508; }
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

    // ── Geração e download do arquivo Excel ─────────────────────────────────────
    // Chamado a partir de um clique do usuário (gesto real), evitando o bloqueio
    // de downloads automáticos que o navegador impõe a downloads sem interação.

    function gerarEbaixarExcel(dados) {
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

    // ── Estados do botão ────────────────────────────────────────────────────────

    function configurarBotaoExportar() {
        const btn = document.getElementById('btn-exportar-excel');
        if (!btn) return;
        btn.disabled = false;
        btn.className = '';
        btn.textContent = 'Exportar Excel';
        btn.onclick = iniciarExportacao;
    }

    function configurarBotaoBaixar(qtd) {
        const btn = document.getElementById('btn-exportar-excel');
        if (!btn) return;
        btn.disabled = false;
        btn.className = 'pronto';
        btn.textContent = `⬇ Baixar planilha (${qtd} registros)`;
        btn.onclick = function () {
            try {
                const dados = JSON.parse(sessionStorage.getItem(KEY_DADOS) || '[]');
                gerarEbaixarExcel(dados);
                atualizarStatus(`✓ ${dados.length} registros exportados.`);
                sessionStorage.removeItem(KEY_PRONTO);
                sessionStorage.removeItem(KEY_DADOS);
                configurarBotaoExportar();
            } catch (err) {
                atualizarStatus(`Erro ao gerar planilha: ${err.message}`);
                console.error('[Exportar Projudi]', err);
            }
        };
    }

    // ── Lógica principal ────────────────────────────────────────────────────────
    //
    // A paginação do Projudi faz RELOAD COMPLETO da página a cada troca de página.
    // Por isso usamos sessionStorage para acumular dados entre os reloads.
    // Ao terminar a coleta NÃO baixamos automaticamente (o navegador bloqueia
    // downloads sem gesto do usuário); em vez disso o botão vira "Baixar planilha"
    // e o usuário clica para disparar o download.

    function iniciarExportacao() {
        sessionStorage.setItem(KEY_DADOS, JSON.stringify([]));
        sessionStorage.setItem(KEY_RODANDO, '1');
        sessionStorage.removeItem(KEY_PRONTO);
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
        const porPagina = dadosPagina.length || 20;
        const totalPags = total ? Math.ceil(total / porPagina) : '?';

        atualizarStatus(`Coletando página ${pagina} de ${totalPags} — ${dadosAcumulados.length} registros até agora...`);

        if (temProximaPagina()) {
            // Clica na próxima página — isso recarrega a página;
            // o script retoma em continuarExportacao() no próximo load.
            document.querySelector('a.arrowNextOn').click();
        } else {
            // Última página: marca como pronto e deixa o usuário baixar com um clique
            sessionStorage.removeItem(KEY_RODANDO);
            sessionStorage.setItem(KEY_PRONTO, '1');
            atualizarStatus(`Coleta concluída: ${dadosAcumulados.length} registros. Clique para baixar a planilha.`);
            configurarBotaoBaixar(dadosAcumulados.length);
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
        btn.type = 'button';
        btn.title = 'Exporta todos os registros de todas as páginas para .xlsx';
        buttonBar.appendChild(btn);

        const status = document.createElement('span');
        status.id = 'exportar-status';
        buttonBar.appendChild(status);

        if (sessionStorage.getItem(KEY_RODANDO) === '1') {
            // Exportação em andamento, retomada após reload de página
            continuarExportacao();
        } else if (sessionStorage.getItem(KEY_PRONTO) === '1') {
            // Coleta terminou em um ciclo anterior mas o arquivo ainda não foi baixado
            const qtd = JSON.parse(sessionStorage.getItem(KEY_DADOS) || '[]').length;
            atualizarStatus(`Coleta concluída: ${qtd} registros. Clique para baixar a planilha.`);
            configurarBotaoBaixar(qtd);
        } else {
            configurarBotaoExportar();
        }
    }

    if (document.readyState === 'loading') {
        document.addEventListener('DOMContentLoaded', injetarBotao);
    } else {
        injetarBotao();
    }
})();
