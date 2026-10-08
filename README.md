# Exportar Dados Projudi para Excel ou Relatório PDF

Userscript para [Tampermonkey](https://www.tampermonkey.net/) que automatiza a coleta de dados no sistema [Projudi (TJPR)](https://projudi2.tjpr.jus.br/) e exporta relatórios em planilha Excel (`.xlsx`) ou PDF — com resumo, gráficos e a lista detalhada dos processos — servindo de base para o **Relatório para Correição Ordinária**.

Exige login de servidor ou Corregedoria, assim como vinculação à unidade judicial cujos dados pretende se exportar.

## Relatórios suportados

Os relatórios aparecem agrupados no painel de automação, em duas abas: **Cível-Geral** e **Crime**.

### Cartório

- **Estatísticas gerais** — Ativos por Classe Processual · Suspensos por Prazo Indeterminado · Suspensos com Prazo · Em Instância Recursal
- **Pendências** — Juntadas · Retorno de Conclusos · Processos Paralisados · Remessas em Aberto · Processos Remetidos (somados ao indicador de Remessas)
- **Mandados** — Aguardando Análise de Retorno · Aguardando Distribuição ao Oficial de Justiça · Pendentes de Cumprimento · Aguardando Análise de Decurso de Prazo
- **Audiências** — Pendentes · Designadas · Realizadas
- **Outros** — Tempo Médio de Cumprimento · Outros Cumprimentos (painel de contadores da Mesa do Magistrado) · Processos Arquivados com Saldo

### Gabinete

- **Conclusões** — processos aguardando conclusão, com gráfico por magistrado(a) e opção de **PDF por Juiz**

### VIJ — Seção Cível

- Crianças/Adolescentes Acolhidos
- Habilitações para Adoção (Aguardando Oportuna Indicação)

### VIJ — Seção Infracional

- Adolescentes Internados

### Família

- Averiguação de Paternidade (Processos Ativos)
- Prisões — Alimentos

Nos relatórios de Acolhidos, Internados e Prisões — Alimentos, o resumo conta **pessoas distintas** (Nome + Data de Nascimento) além do número de processos, e alerta quando há pessoas sem data de nascimento cadastrada.

### Crime (aba própria no painel)

- Apreensões Pendentes · Bens Pendentes de Cadastro no SNGB
- Mandados de Prisão a Regularizar · Alvarás de Soltura a Regularizar
- Prisões (Réus Presos) · Reavaliação de Prisão Provisória (art. 316, CPP)
- Cumprimento de Medidas · Medidas Alternativas em Atraso · Monitoração Eletrônica Expirada
- Transação Penal e Suspensão Condicional do Processo (ativas — e as variantes com processo ativo)
- Prescrições · Feitos Sem Infração Penal
- Feitos com Réu sem RG/IIPR · Feitos com Réu sem CPF/CNPJ · Campos Obrigatórios Pendentes da Parte (Violência Doméstica)
- **Tribunal do Júri** — Listas de Jurados (últimos 5 anos)

## Funcionalidades

- Percorre automaticamente todas as páginas de resultado de cada relatório
- Acumula dados de **múltiplas Atuações/competências** antes de exportar
- Exporta para `.xlsx` via [SheetJS](https://sheetjs.com/) ou para PDF via [jsPDF](https://github.com/parallax/jsPDF) + [jspdf-autotable](https://github.com/simonbengtsson/jsPDF-AutoTable), com resumo (indicadores e gráficos) e tabela detalhada
- **Relatório conjunto em PDF**, em dois formatos:
  - **Relatório PDF** — capa unificada e resumo de cada relatório coletado (indicadores e tabela comparativa por competência), com índice clicável e marcadores de navegação
  - **Tabelas Discriminadas** — só as listas detalhadas de processos de todos os relatórios coletados
  - Antes de gerar, é possível escolher quais atribuições/competências entram no PDF
- **Painel de automação**: navega sozinho entre os relatórios selecionados, preenche filtros, pesquisa e coleta todas as páginas de cada um, sem intervenção manual
- **Automação em várias unidades**: na tela "Selecione a Área de Atuação", marque várias unidades e os relatórios desejados — o script roda a automação em cada unidade, uma após a outra, trocando de atuação sozinho
- **Exportar/importar preferências** (`.json`): leva os relatórios marcados no painel para outro computador ou os preserva antes de reinstalar o script
- Persiste os dados coletados no `localStorage`, tolerando recarregamentos de página durante a paginação automática
- Detecta coletas travadas (sem atividade por mais de 6 minutos) e descarta apenas a flag de execução, preservando os dados já coletados; o botão **⏭ Pular extração atual** permite seguir adiante manualmente (o relatório pulado consta no PDF como interrompido por erro)
- Se o armazenamento do navegador encher durante a coleta, a automação para com um aviso — exporte o que já foi coletado e use **Limpar** antes de continuar
- **Log de execução** no painel, com opção de baixar o log completo (`.txt`) para investigar problemas
- **Aviso de atualização**: a cada 30 minutos, na página inicial, o script verifica se há versão nova publicada e mostra um aviso com o botão **Atualizar agora**

## Instalação

1. Instale a extensão [Tampermonkey](https://www.tampermonkey.net/) no seu navegador
2. Abra o arquivo [`relatorio_projudi.user.js`](./relatorio_projudi.user.js) e clique em **Raw** — o Tampermonkey reconhece a extensão e oferece a instalação automaticamente

Ou, manualmente: painel do Tampermonkey → **Criar novo script** → cole o conteúdo de `relatorio_projudi.user.js` → salve (Ctrl+S).

### Atualização

Quando houver versão nova, o script mostra um aviso na página inicial do Projudi; clique em **Atualizar agora** e confirme na tela do Tampermonkey. Também é possível reabrir a URL Raw a qualquer momento — o Tampermonkey substitui a versão antiga.

## Como usar

### Relatório individual

1. Acesse a tela do relatório desejado no Projudi — em muitos relatórios há o botão **Preencher e Pesquisar**, que ajusta os filtros e pesquisa por você; ou pesquise normalmente
2. Clique em **Extrair** (ou **Coletar**, conforme o relatório) — o script percorre todas as páginas de resultado automaticamente; para acumular outra atuação, use **Extrair mais...**
3. Clique em **⬇ Baixar** para gerar a planilha Excel, ou em **⬇ Baixar PDF** para gerar o relatório em PDF (em Conclusões há também **PDF por Juiz**)
4. Use **Limpar** para apagar os dados acumulados desse relatório

### Automação (relatório conjunto entre vários relatórios/competências)

1. Abra a página inicial do Projudi — o painel **"Automação de relatórios"** aparece nas páginas com o menu principal
2. Escolha a aba (**Cível-Geral** ou **Crime**), marque os relatórios desejados e clique em **▶ Iniciar** — a automação coleta tudo o que estiver marcado, nas duas abas
3. O script navega, preenche filtros, pesquisa e coleta cada relatório marcado, na sequência
   - **⏸ Parar** pausa a extração; tudo o que já foi coletado fica gravado
   - **▶ Continuar** retoma de onde parou, mantendo os dados já coletados (se, durante a pausa, você sair da página que estava sendo coletada, a etapa em andamento é refeita do início, sem duplicar dados)
   - **⏭ Pular extração atual** segue para o próximo relatório em caso de travamento
4. Para reunir mais de uma competência: troque de atuação no Projudi e rode a automação de novo — os dados de cada rodada são **acumulados**, não substituídos (ou use a automação em várias unidades, abaixo)
5. Ao concluir, clique em **⬇ Relatório PDF** ou **⬇ Tabelas Discriminadas** (ou baixe o Excel de cada relatório individualmente)
6. Use **Limpar** no painel para apagar tudo o que foi acumulado

### Automação em várias unidades

1. Na tela **"Selecione a Área de Atuação"** (no login ou pelo botão de alterar atuação), marque as unidades desejadas na árvore
2. No painel **"Automação em várias unidades"**, escolha os relatórios (há atalhos por perfil: Cível, Família, Infância, Crime, Todos)
3. Clique em **▶ Rodar automação nas unidades marcadas** — o script entra em cada unidade, roda a automação completa e passa para a próxima
4. Ao final, gere o PDF conjunto pelo painel da página inicial

## Requisitos

- Tampermonkey instalado no navegador (Chrome, Firefox, Edge, etc.)
- Acesso ao sistema Projudi do TJPR (`projudi2.tjpr.jus.br`)

## Dependências externas

Carregadas via `@require` pelo Tampermonkey:

- [SheetJS (xlsx)](https://cdn.sheetjs.com/) — exportação para Excel
- [jsPDF](https://cdnjs.com/libraries/jspdf) — geração de PDF
- [jspdf-autotable](https://cdnjs.com/libraries/jspdf-autotable) — tabelas no PDF
