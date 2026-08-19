# Exportar Conclusões Projudi para Excel

Userscript para [Tampermonkey](https://www.tampermonkey.net/) que automatiza a coleta de dados no sistema [Projudi (TJPR)](https://projudi2.tjpr.jus.br/) e exporta relatórios em planilha Excel (`.xlsx`) ou PDF — com resumo, gráficos e a lista detalhada dos processos.

📘 Manual do usuário (instalação e uso passo a passo): [`manual_usuario.pdf`](./manual_usuario.pdf) / [`manual_usuario.docx`](./manual_usuario.docx)

## Relatórios suportados

- **Conclusões** — processos aguardando conclusão
- **Retorno de Processos Conclusos**
- **Juntadas** — juntadas pendentes de análise
- **Tempo Médio de Cumprimento**
- **Processos Paralisados**
- **Remessas em Aberto**

## Funcionalidades

- Percorre automaticamente todas as páginas de resultado de cada relatório
- Acumula dados de **múltiplas Atuações/competências** antes de exportar
- Exporta para `.xlsx` via [SheetJS](https://sheetjs.com/) ou para PDF via [jsPDF](https://github.com/parallax/jsPDF) + [jspdf-autotable](https://github.com/simonbengtsson/jsPDF-AutoTable), com resumo (KPIs e gráficos) e tabela detalhada
- Opção de gerar o PDF **só com o resumo**, sem a tabela discriminada de processos
- **Relatório conjunto**: reúne vários relatórios num único PDF, com capa, índice clicável e marcadores de navegação — a capa só aparece quando os dados reúnem mais de uma competência; com uma só, sai direto como relatório único
- **Painel de automação**: navega sozinho entre os relatórios selecionados, preenche filtros, pesquisa e coleta todas as páginas de cada um, sem intervenção manual
- Persiste os dados coletados no `localStorage`, tolerando recarregamentos de página durante a paginação automática
- Detecta coletas travadas (sem atividade por mais de 2 minutos) e descarta apenas a flag de execução, preservando os dados já coletados

## Instalação

1. Instale a extensão [Tampermonkey](https://www.tampermonkey.net/) no seu navegador
2. Acesse o painel do Tampermonkey → **Criar novo script**
3. Copie e cole o conteúdo de [`conclusoes_projudi.user.js`](./conclusoes_projudi.user.js)
4. Salve (Ctrl+S)

Ou clique em **Raw** no arquivo `.user.js` — o Tampermonkey reconhece a extensão e oferece a instalação automaticamente.

Ao publicar uma nova versão, repita os passos acima (ou reabra a URL Raw) — o Tampermonkey substitui a versão antiga automaticamente.

## Como usar

### Relatório individual

1. Acesse a tela do relatório desejado no Projudi e realize a pesquisa normalmente
2. Clique em **Extrair** (ou **Coletar**, conforme o relatório) — o script percorre todas as páginas de resultado automaticamente; para continuar depois, use **Extrair mais...**
3. Clique em **⬇ Baixar** para gerar a planilha Excel, ou em **⬇ Baixar PDF** para gerar o relatório em PDF (marque **"Só resumo (sem tabela)"** para omitir a tabela detalhada)
4. Use **Limpar** para apagar os dados acumulados desse relatório

### Automação (relatório conjunto entre vários relatórios/competências)

1. Abra a página inicial do Projudi — o painel **"Automação de relatórios"** aparece nas páginas com o menu principal
2. Marque os relatórios desejados e clique em **▶ Automatizar**
3. O script navega, preenche filtros, pesquisa e coleta cada relatório marcado, na sequência
4. Para reunir mais de uma competência: troque de atuação no Projudi e rode a automação de novo — os dados de cada rodada são **acumulados**, não substituídos
5. Ao concluir, clique em **⬇ PDF conjunto** (ou baixe o Excel de cada relatório individualmente)
6. Use **Limpar** no painel para apagar tudo o que foi acumulado

Veja o [manual do usuário](./manual_usuario.pdf) para o passo a passo completo, opções do PDF e solução de problemas comuns.

## Requisitos

- Tampermonkey instalado no navegador (Chrome, Firefox, Edge, etc.)
- Acesso ao sistema Projudi do TJPR (`projudi2.tjpr.jus.br`)

## Dependências externas

Carregadas via `@require` pelo Tampermonkey:

- [SheetJS (xlsx)](https://cdn.sheetjs.com/) — exportação para Excel
- [jsPDF](https://cdnjs.com/libraries/jspdf) — geração de PDF
- [jspdf-autotable](https://cdnjs.com/libraries/jspdf-autotable) — tabelas no PDF

## Autor

[@rcpleme2](https://github.com/rcpleme2)
