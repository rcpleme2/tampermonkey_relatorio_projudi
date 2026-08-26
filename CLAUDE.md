# CLAUDE.md

Orientações para trabalhar neste repositório com Claude Code — arquitetura do userscript,
convenções já estabelecidas e um roteiro para implementar novos relatórios/ferramentas
seguindo os mesmos padrões.

## O que é o projeto

Userscript Tampermonkey (`conclusoes_projudi.user.js`, arquivo único, ~5500+ linhas) que
automatiza a extração de dados do sistema Projudi (TJPR) e gera relatórios em Excel
(`.xlsx`, via SheetJS) e PDF (via jsPDF + jspdf-autotable). Cobre relatórios de Cartório
(Conclusões, Retorno, Juntadas, Paralisados, Remessas, Suspensos, Tempo Médio, Audiências,
Apreensões, Outros Cumprimentos) e de Gabinete (Conclusões por magistrado).

Documentação voltada ao usuário final: `README.md` e `manual_usuario.pdf/docx`. Este
arquivo é sobre como **implementar** no código, não como usar o script.

## Estrutura do arquivo — ATENÇÃO a linhas gigantes

`conclusoes_projudi.user.js` tem 3 linhas enormes (`FONTE_PUBLIC_SANS_REGULAR/BOLD/ITALIC`
— fontes embutidas em base64, ~75 mil caracteres cada) perto da função `novoDocPDF()`.

**Nunca leia essas linhas inteiras.** A ferramenta `Read` explode em tokens se o
offset/limit passar por elas (mesmo pedindo poucas linhas — o tokenizer processa o
conteúdo antes de aplicar o limite). Para ler qualquer trecho do arquivo:

- Prefira `Grep` com contexto pequeno (`-A`/`-B`/`-C`) para achar funções por nome.
- Para ler um trecho por número de linha, use `sed -n 'X,Yp' conclusoes_projudi.user.js`
  via Bash em vez de `Read` com offset/limit, e confira antes (`grep -n` ou `awk
  '{print length, NR}' | sort -rn | head`) que o intervalo não cruza essas 3 linhas.
- `node --check conclusoes_projudi.user.js` funciona normalmente (não precisa ler o
  arquivo pra isso).

## Arquitetura — como um relatório "normal" (paginado) funciona

A maioria dos relatórios segue este pipeline, todo dentro do mesmo IIFE do userscript:

1. **`CFG_<RELATORIO>`** — objeto de configuração central: `prefixo` (chave no
   `localStorage`), `detecta(cab)` (regex sobre o cabeçalho da `table.resultTable` que
   reconhece a tela), `minTds`, `usaAtuacao`, `nomeArquivo`, `rotulos` (textos dos
   botões), `cabecalhos`/`larguras`/`extrai`/`linha` (colunas do Excel) e `pdf` (título,
   `atosTitulo`, `agingTitulo`, `dataCampo`, `processoCampo`, `tipoCampo`,
   `distribuicoes` — gráficos de barra por campo —, `colunas` da tabela discriminada, e
   flags opcionais como `mediaLabel`, `ordenarPrioritarioPrimeiro`, `semPrioridade`,
   `agruparPor`/`ordemGrupos`).
2. **`REPORTS_AUTOMACAO`** — array com um item por relatório: `{ key, cfg, navAlvo,
   rotulo, curto, categoriaEspecifica, precisaPreencher }`. É o que popula o painel de
   automação e a fila (`lerFilaAutomacao`).
3. **`navegarMenu(alvo)`** — acha o link/aba de menu certo para o relatório e navega
   (clique real, não `location.href`, quando o elemento não expõe `href`). Cada
   relatório novo precisa de um `else if` aqui.
4. **`detectarConfig()`** — dado o cabeçalho de uma `table.resultTable` na tela atual
   (ou outro sinal de página, para telas atípicas), decide qual `CFG_*` está ativo.
5. **Preenchimento + pesquisa** — função dedicada (ex. `preencherEPesquisarApreensoes`)
   que ajusta os filtros do formulário do Projudi e clica em Pesquisar.
6. **Coleta paginada** — `criarColetor(cfg)` genérico: lê `table.resultTable tbody tr`
   da página atual via `cfg.extrai(tds, atuacao)`, salva incrementalmente em
   `localStorage` (`prefixo+'pagina_N'`, `prefixo+'num_paginas'`, `prefixo+'coletado'`),
   avança para a próxima página do Projudi, repete. Tolera reload de página.
7. **PDF individual** — `gerarPDF<Relatorio>(dados)`, normalmente delegando para
   `montarResumoGenerico`/`montarTabelaGenerico` (funções genéricas movidas por
   `cfg.pdf`) quando os dados são "um registro por processo". Quando os dados **não**
   são por processo (ex. painel de contadores agregado), escreva funções de PDF
   dedicadas — não force o encaixe genérico (ver seção "Outros Cumprimentos" abaixo
   como referência).
8. **Integração na capa unificada** (`gerarPDFConjunto`/`desenharBlocoCartorioUnificado`
   para Cartório, `desenharBlocoDominio` para Gabinete) — uma linha na tabela do
   Cartório com `{nome, indicador, detalhamento, situacaoLabel, corTexto, semSituacao,
   cfgOriginal}`, linkada por página à seção detalhada via `_rect`/`doc.link` (ver
   "PASSO 4" dentro de `gerarPDFConjunto`).
9. **Botões manuais** — `injetarBotoes()` injeta Extrair/Baixar/Limpar na
   `table.buttonBar` da tela (ou noutro ponto, se a tela não tiver `buttonBar` — ver
   Outros Cumprimentos).

### Convenção Cartório vs. Gabinete

Estabelecida numa sessão de `/design`: **todo relatório novo é Cartório**, a não ser que
seja especificamente sobre conclusões por magistrado — isso fica em Gabinete. Não
pergunte de novo ao usuário sobre isso; siga essa regra por padrão.

## Quando o relatório NÃO é uma lista paginada por processo

Nem todo relatório do Projudi é uma busca com paginação. **Outros Cumprimentos** (Mesa do
Magistrado) é um painel de contadores agregados por tipo, numa única tela, sem paginação
e sem botão de pesquisar. Para casos assim:

- Não force `criarColetor()`/`montarResumoGenerico()`/`montarTabelaGenerico()` — eles
  assumem "um registro por processo, com data e número de processo". Escreva uma função
  de coleta dedicada (ex. `coletarOutrosCumprimentosAgora()`) e funções de PDF dedicadas
  (`montarResumo<X>`/`montarTabela<X>`), reaproveitando os *helpers visuais* (`COR.*`,
  fonte `PublicSans`, `desenharCard`, `tituloSecao`, `desenharRodape`, `doc.autoTable`
  com `theme:'grid'` e os mesmos estilos) para manter a identidade visual.
- Salve os dados coletados no MESMO formato (`prefixo+'pagina_0'` como array JSON,
  `num_paginas`, `coletado`) para que `lerDadosDe`/`foiColetado`/a capa unificada
  continuem funcionando sem mudança.
- Ainda assim registre em `REPORTS_AUTOMACAO`, `navegarMenu`, `detectarConfig` (ou uma
  função de detecção própria tipo `paginaOutrosCumprimentos()`, se a detecção genérica
  por cabeçalho de `table.resultTable` não servir).

## Armadilhas já descobertas (não repita)

- **Clique sintético não é clique de verdade.** O JS de página do Projudi só reage a um
  evento de clique real do usuário em certos elementos (checkbox "Todas" em Audiências
  Designadas, aba "Outros Cumprimentos" sem `href`). Fazer `elemento.checked = true` +
  `dispatchEvent(new Event('change'))` NÃO dispara o cascateamento/navegação esperado —
  é preciso `elemento.click()` de verdade. Quando um fluxo de preenchimento parar de
  funcionar silenciosamente, suspeite disso primeiro.
- **Tabelas carregadas de forma assíncrona.** Nem toda `table.resultTable` está pronta
  assim que a página "termina" de carregar — Outros Cumprimentos preenche sua tabela
  principal via AJAX depois do HTML inicial. Extrair cedo demais pega dados parciais.
  Padrão de correção: comparar uma "assinatura" do DOM (nº de linhas + soma dos
  contadores) em leituras sucessivas com `setTimeout`, só extrair quando estabilizar (2
  leituras iguais seguidas), com um teto de tempo como salvaguarda (ver
  `aguardarOutrosCumprimentosProntoEExtrair`).
- **`Event` não existe em sandbox `vm` puro.** Testes que chamam código com `new
  Event(...)` dentro de `vm.createContext()` precisam de um polyfill de `Event` passado
  explicitamente no objeto de contexto — o `Event` global do Node não vaza pra dentro do
  sandbox.
- **`:scope > td` em vez de `td`** ao contar colunas de uma `<tr>` — algumas células têm
  tabelas aninhadas (ex. coluna "Partes" de Audiências), e `td` puro pegaria os `<td>` de
  dentro também, bagunçando os índices.
- **Truncamento de texto em cards de PDF.** Cards de KPI têm largura fixa; texto que não
  cabe corta feio. Use `textoTruncadoParaLargura(doc, texto, larguraMax)` quando um
  resumo curto é aceitável, ou `medirAlturaCardLista`/`desenharCardLista` (altura
  dinâmica, quebra de linha, nunca corta) quando o conteúdo precisa aparecer por inteiro
  (ex. lista de números de processo).

## Fluxo para implementar um novo relatório

1. Peça (ou receba) uma amostra real da tela do Projudi — idealmente um `.mhtml` salvo
   com resultados já carregados, e/ou um trace do puppeteer de como chegar até lá
   (`navegarMenu`). Decodifique `.mhtml` com `quopri`/Python se precisar ler a marcação
   sem os `=3D`/quebras de linha do quoted-printable.
2. Se o volume de dados a interpretar for grande (HTML de uma tela real, JSON de
   referência), salve um recorte já processado no `scratchpad/` (JSON com os campos
   relevantes) para servir de fixture de teste — evita reprocessar o `.mhtml` inteiro
   depois.
3. Se o escopo do relatório for ambíguo (ex. "relate os pendentes" pode significar coisas
   diferentes dependendo da tela), pergunte ao usuário ANTES de implementar — mudar de
   abordagem depois de pronto custa mais caro que uma pergunta.
4. Implemente seguindo os pontos de integração da seção "Arquitetura" acima. Prefira
   reaproveitar o genérico; só desvie dele quando o formato dos dados genuinamente não
   servir (documente a decisão no comentário do código e na mensagem de commit).
5. `node --check conclusoes_projudi.user.js` depois de qualquer edição.
6. Escreva/atualize um teste em `scratchpad/teste_<relatorio>.js` (ver seção Testes) e
   rode a suíte de regressão existente (pelo menos `verify_conclusoes.js`,
   `teste_cartorio_unificado.js`, `teste_zero_e_links.js`) para garantir que a mudança
   não quebrou outro relatório.
7. Bump de `@version` no cabeçalho `UserScript` (regra: **sempre** que o arquivo mudar,
   mesmo em ajustes pequenos — o Tampermonkey usa isso pra saber que há atualização).
8. Commit em português, mensagem explicando o **porquê** da mudança (não só o quê),
   rodapé padrão do projeto (ver seção Git).

## Testes

Não há framework de teste formal — os testes vivem em `scratchpad/` (fora do repo git,
específico da sessão) como scripts Node standalone usando o módulo `vm`:

- Lê o `.user.js` real, remove o cabeçalho `// ==UserScript== ... ==/UserScript==`, e
  injeta antes do `})();` final um `globalThis.__T = { ...símbolos internos... }` — só
  assim funções `function`-scoped dentro do IIFE ficam acessíveis para o teste.
- Cria um `context` mínimo (`document`/`window.localStorage`/`console`/`Blob`/`URL` com
  `createObjectURL` capturando o blob gerado/`GM_addStyle`/`GM_download`/timers) e roda
  com `vm.createContext` + `vm.runInContext`.
- Usa `jsPDF`+`jspdf-autotable` REAIS (de `scratchpad/v251check/node_modules/`) para
  gerar PDFs de verdade a partir de dados sintéticos — não é mock, então pega erros reais
  de `autoTable`/paginação.
- Para telas com DOM mais elaborado (múltiplas tabelas, spans com `id`), construa um DOM
  sintético mínimo com só os métodos usados pelo script real
  (`querySelector`/`querySelectorAll`/`textContent`/`nextElementSibling` etc.) — não
  tem jsdom disponível no ambiente.
- Para funções assíncronas (`setTimeout`-based, ver a armadilha de tabelas assíncronas
  acima), teste o comportamento de verdade: mude o DOM sintético num timer real e
  confirme que a função esperou e pegou o estado final, não o inicial.
- Padrão de nomes: `scratchpad/teste_<relatorio ou funcionalidade>.js`. Gera PDFs de
  saída em `scratchpad/*.pdf` — úteis para inspeção visual manual, mas não precisam ser
  commitados (o `scratchpad/` já está fora do repo).

## Git — convenções deste projeto

- Um branch por funcionalidade/relatório (ex. `apreensoes`, `OutrosCumprimentos`), nunca
  commitar direto em `main` a não ser que o usuário peça explicitamente ("corrija direto
  na main").
- PR e merge só quando o usuário pedir explicitamente ("abrir PR e mergear em main") —
  não presuma.
- Mensagens de commit em português, focadas no *porquê*, com o rodapé:
  ```
  Co-Authored-By: Claude Sonnet 5 <noreply@anthropic.com>
  Claude-Session: https://claude.ai/code/session_XXXXXXXX
  ```
- `git push origin <remote antigo>` continua funcionando mesmo o GitHub avisando que o
  repositório mudou de nome (`tampermonkey_conclusoes_projudi` →
  `tampermonkey_relatorio_projudi`) — o push é redirecionado com sucesso; não é um erro.
- Para tarefas grandes/paralelizáveis (ex. implementar um relatório inteiro do zero
  seguindo um padrão já mapeado), delegar para um agente em background com uma
  especificação bem detalhada (contexto do domínio, decisões de design já tomadas,
  trechos de referência) costuma ser mais eficiente do que implementar linha a linha na
  conversa principal — mas sempre revise o diff e rode os testes depois.

## Coisas para NUNCA fazer sem confirmação explícita

- Adicionar dependências novas via `@require` sem necessidade clara.
- Mudar a regra de negócio de um relatório já em produção (ex. quais campos somam num
  total) sem confirmar com o usuário — esses números às vezes vêm de decisões de domínio
  jurídico não óbvias pelo código (ex. "Com Urgência" é subconjunto, não fila própria).
- Remover funcionalidade "porque parece não usada" sem grep completo — o arquivo é
  grande e interligado (capa unificada, automação, botões manuais costumam referenciar a
  mesma `CFG_*` de lugares diferentes).
