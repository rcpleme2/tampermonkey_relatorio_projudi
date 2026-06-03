# Exportar Conclusões Projudi para Excel

Userscript para [Tampermonkey](https://www.tampermonkey.net/) que coleta as conclusões judiciais do sistema [Projudi (TJPR)](https://projudi2.tjpr.jus.br/) e exporta tudo em uma planilha Excel (`.xlsx`).

## Funcionalidades

- Percorre automaticamente todas as páginas da tela de conclusões
- Acumula dados de **múltiplas Atuações** antes de exportar
- Exporta para `.xlsx` com colunas pré-formatadas via [SheetJS](https://sheetjs.com/)
- Persiste os dados coletados no `localStorage`, tolerando recarregamentos de página durante a paginação automática
- Detecta coletas travadas (sem atividade por mais de 2 minutos) e descarta apenas a flag de execução, preservando os dados já coletados

## Colunas exportadas

| Coluna | Descrição |
|---|---|
| Atuação | Área de atuação selecionada no momento da coleta |
| Dt. Remessa | Data de remessa da conclusão |
| Processo | Número do processo |
| Classe | Classe processual |
| Seq. | Sequência |
| Tipo de Conclusão | Tipo da conclusão |
| Privativa | Indicador de privatividade |
| Responsável | Nome do responsável |
| Pré-análise | Informação de pré-análise |
| Agrupador | Agrupador (quando disponível) |

## Instalação

1. Instale a extensão [Tampermonkey](https://www.tampermonkey.net/) no seu navegador
2. Acesse o painel do Tampermonkey → **Criar novo script**
3. Copie e cole o conteúdo de [`conclusoes_projudi.user.js`](./conclusoes_projudi.user.js)
4. Salve (Ctrl+S)

Ou clique em **Raw** no arquivo `.user.js` — o Tampermonkey reconhece a extensão e oferece a instalação automaticamente.

## Como usar

1. Acesse a tela de conclusões no Projudi (`/projudi/processo/conclusao.do`)
2. Selecione a **Atuação** desejada
3. Clique em **Coletar esta Atuação** — o script percorre todas as páginas automaticamente
4. Para acumular mais de uma atuação: troque de atuação no sistema e clique em **Coletar mais uma Atuação**
5. Quando tiver coletado tudo, clique em **⬇ Baixar planilha** para gerar o arquivo Excel
6. Use **Limpar** para apagar os dados acumulados e começar uma nova coleta

## Requisitos

- Tampermonkey instalado no navegador (Chrome, Firefox, Edge, etc.)
- Acesso ao sistema Projudi do TJPR (`projudi2.tjpr.jus.br`)

## Dependências externas

- [SheetJS (xlsx)](https://cdn.sheetjs.com/) — carregado via `@require` pelo Tampermonkey

## Autor

[@rcpleme2](https://github.com/rcpleme2)
