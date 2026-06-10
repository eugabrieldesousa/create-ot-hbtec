# Performance React - Rounds 2 e 3

Data: 2026-06-10

## Secao 1 - O que foi feito

| Tarefa | Status | Arquivos modificados | Observacao |
| --- | --- | --- | --- |
| A1. Code splitting do bundle principal | Concluido | `src/App.tsx`, `src/styles.css`, `vite.config.ts`, `vite.config.js` | `docxPreview`, `docxExport` e `docxImport` sairam do grafo inicial. `docx` e `mammoth` ficaram em `vendor-docx`. |
| A2. Memoizacao de PermissionGroupEditor | Concluido | `src/App.tsx` | `PermissionGroupEditor` recebeu `memo` com comparador por dados relevantes; handlers de macro/micro foram estabilizados no `App`. `PermissionBlockGroup` e `PermissionBlockEditor` ja estavam memoizados no estado atual do codigo. |
| A3. Keys estaveis no DocxPreview | Concluido | `src/docxPreview.tsx` | Keys deixaram de usar indice puro; listas sem id proprio usam tipo, posicao contextual e prefixo de texto. |
| A4. Auditar useEffect com dependencias instaveis | Concluido | `src/App.tsx` | `useActiveOutlineTargetId` deixou de fechar sobre `targetIds`; effect de hidratacao inicial documentado como montagem intencional. |
| A5. why-did-you-render | Parcial | `package.json`, `package-lock.json` | Instalado `@welldone-software/why-did-you-render@8.0.3` por compatibilidade com React 18. Codigo de producao nao importa WDYR. Fluxo no navegador nao foi executado nesta sessao. |
| B1. Virtualizacao | Bloqueado | Nenhum | Requer baseline real com React DevTools Profiler e documento com 30+ atividades. Nao foi adicionada dependencia de producao sem medicao. |
| B2. Web Vitals e Long Tasks | Bloqueado | Nenhum | Requer Chrome Performance tab e execucao interativa. `web-vitals` nao foi instalado para evitar dependencia de producao sem coleta. |
| B3. Profiling de memoria | Bloqueado | Nenhum | Requer Heap Snapshots no Chrome DevTools. Auditoria de codigo verificou cleanups existentes em timers/listeners tocados. |
| B4. Inputs restantes | Concluido sem diff | `src/App.tsx` auditado | Inputs de texto pesados ja usam `useBufferedText`, `BufferedTextInput`, `BufferedTextarea` ou uncontrolled com commit explicito. |
| B5. Arquitetura de estado | Documentado | Nenhum | Mapa de estado e proximas 3 mudancas de maior impacto documentados abaixo. |
| B6. Relatorio final | Concluido | `PERFORMANCE_ROUND2_3.md` | Este arquivo. |

Possivel efeito colateral: o preview DOCX agora carrega sob demanda. Na primeira abertura da aba de preview, pode haver um micro atraso enquanto o chunk `docxPreview` e baixado/carregado; o `PreviewSkeleton` mantem a area estavel para evitar layout shift.

## Secao 2 - Metricas comparativas

| Metrica | Antes Round 2 | Depois Round 2/3 | Status |
| --- | ---: | ---: | --- |
| Chunk principal JS minificado | 1,411.57 kB | 509.96 kB | Melhorou 901.61 kB |
| Chunk principal JS gzip | 398.96 kB | 153.70 kB | Melhorou 245.26 kB |
| Chunk `docxPreview` | No principal | 2.74 kB / 1.01 kB gzip | Isolado |
| Chunk `docxExport` | No principal | 12.11 kB / 4.32 kB gzip | Isolado |
| Chunk `docxImport` | No principal | 15.93 kB / 5.23 kB gzip | Isolado |
| Chunk `vendor-docx` | No principal | 870.05 kB / 235.89 kB gzip | Isolado sob demanda |
| INP digitacao TEA | Nao medido | Nao medido | Requer navegador |
| INP collapse/expand | Nao medido | Nao medido | Requer navegador |
| Long Tasks >50 ms | Nao medido | Nao medido | Requer Chrome Performance |
| FPS minimo collapse/expand | Nao medido | Nao medido | Requer Chrome Performance |
| CLS no lazy preview | Nao medido | Esperado 0 | Skeleton ocupa a area do preview, mas falta confirmar no Chrome |

## Secao 3 - Bundle size

Saida final de `npm run build`:

```text
dist/index.html                                      0.50 kB | gzip:   0.31 kB
dist/assets/inter-latin-400-normal-C38fXH4l.woff2  23.66 kB
dist/assets/inter-latin-500-normal-Cerq10X2.woff2  24.27 kB
dist/assets/inter-latin-700-normal-Yt3aPRUw.woff2  24.36 kB
dist/assets/inter-latin-600-normal-LgqL8muc.woff2  24.45 kB
dist/assets/inter-latin-400-normal-CyCys3Eg.woff   30.70 kB
dist/assets/inter-latin-600-normal-CiBQ2DWP.woff   31.26 kB
dist/assets/inter-latin-500-normal-BL9OpVg8.woff   31.28 kB
dist/assets/inter-latin-700-normal-BLAVimhd.woff   31.32 kB
dist/assets/index-H_rQz4ju.css                    228.49 kB | gzip:  34.73 kB
dist/assets/docxPreview-CDgTmf2C.js                 2.74 kB | gzip:   1.01 kB
dist/assets/docxExport-BtU0kbSg.js                 12.11 kB | gzip:   4.32 kB
dist/assets/docxImport-ByrNbJYq.js                 15.93 kB | gzip:   5.23 kB
dist/assets/index-C-6X9xYo.js                     509.96 kB | gzip: 153.70 kB
dist/assets/vendor-docx-BIVuQ1E1.js               870.05 kB | gzip: 235.89 kB
```

O aviso de chunk acima de 500 kB permanece porque o limite padrao do Vite e 500 kB: o chunk principal ficou em 509.96 kB e `vendor-docx` em 870.05 kB. O custo de DOCX saiu do carregamento inicial e passa a ser pago somente em import/export DOCX. O chunk principal ficou abaixo da meta solicitada de 600 kB.

## Secao 4 - Testes

Saida final de `npm test`:

```text
Test Files  6 passed (6)
Tests       55 passed (55)
Duration    18.52s
```

Validacoes executadas durante a rodada:

- A1: `npm run build` OK, `npm test` OK.
- A2: `npm run build` OK, `npm test` OK.
- A3: `npm run build` OK, `npm test` OK.
- A4: `npm run build` OK, `npm test` OK.
- A5/final: `npm run build` OK, `npm test` OK.

## Secao 5 - O que NAO foi feito e por que

Virtualizacao com TanStack Virtual nao foi aplicada. A tarefa exige medir primeiro um documento real com 30+ atividades e virtualizar apenas listas acima de 20 itens em uso real. Sem esse dado, adicionar `@tanstack/react-virtual` seria uma dependencia de producao sem prova de ganho e com risco em alturas variaveis de collapse.

Web Vitals nao foi instalado. A tarefa exige coleta de INP/CLS/LCP em fluxos no navegador. Como a coleta nao foi possivel nesta sessao, manter `web-vitals` no bundle de producao nao se justificou.

Chrome Performance, React DevTools Profiler e Heap Snapshots nao foram executados. Essas ferramentas exigem sessao interativa no navegador. O relatorio marca as metricas como nao medidas em vez de inventar numeros.

WDYR foi instalado como devDependency, mas nao foi injetado em `src/main.tsx`. A versao atual do pacote exige React 19; foi escolhida a versao `8.0.3`, que declara peer `react: ^18`. O codigo de producao permanece sem import de WDYR. Para usar localmente, adicionar temporariamente o import dinamico em DEV e remover antes de fechar a investigacao.

## Secao 6 - Proximos passos recomendados

1. Medir com Chrome/React DevTools antes de qualquer nova refatoracao.
   Impacto alto, complexidade baixa. Fluxos: cold load, expand-all TEA, digitar 20 caracteres em bloco TEA, collapse individual.

2. Virtualizar apenas a lista raiz de atividades TEA se commits de expand-all ficarem caros com 30+ atividades reais.
   Impacto medio/alto, complexidade media. Usar `measureElement` por causa de alturas variaveis.

3. Colocar estados de collapse TEA dentro de `TeaWorkspace`.
   Impacto medio, complexidade media. Hoje `collapsedTeaActivities`, `collapsedTeaSubActivities`, `collapsedTeaComposers` e `collapsedTeaContentBlocks` vivem no `App`; mover para baixo reduziria superficie de render do root, mas exige preservar os atalhos de review/outline que abrem paineis.

4. Separar estado OT de testes/permissoes em um subcomponente ou reducer de dominio.
   Impacto medio, complexidade media. Candidatos: `expandedTests`, `collapsedMacros`, `collapsedPermissionBlocks`, `testBlockFilter`.

5. Se o Profiler ainda mostrar root quente, avaliar store com seletores por dominio.
   Impacto alto em documentos grandes, complexidade alta. Evitar antes de ter evidencia, porque a memoizacao atual ja cobre varios caminhos.

## Inventario B4 - Inputs

| Grupo | Categoria | Resultado |
| --- | --- | --- |
| Metadados OT/TEA | Commit buffered | `BufferedTextInput` com `useBufferedText`; nao sobe para o `App` a cada tecla. |
| Objetivo OT, passos em lote, passos individuais | Commit buffered | `BufferedTextarea`/`BufferedTextInput`; commit por debounce/blur/cleanup. |
| Permissoes em lote | Uncontrolled | `defaultValue` + ref + commit explicito, mantido do Round 1. |
| Macro/micro permissao | Commit buffered | Campos usam `BufferedTextInput`; checkboxes sao eventos pontuais. |
| TEA overview/activity intro | Estado local buffered | `useBufferedText` local no `TeaWorkspace` antes de subir ao pai. |
| TEA titulo de atividade/subatividade | Estado local buffered | `useBufferedText` no card. |
| TEA blocos texto/lista | Estado local buffered | `useBufferedText` no editor do bloco. |
| Test title/observations | Estado local buffered | `useBufferedText` no editor de teste. |
| Labels de imagem | Commit buffered | `BufferedTextInput`. |
| Tabs, select, filtros, checks | Controlado reativo | Sao controles de clique/selecao, nao digitacao pesada. |

## Inventario B5 - Estado do App

| Estado | Tamanho tipico | Frequencia | Consumidores principais | Observacao |
| --- | --- | --- | --- | --- |
| `documentKind` | Primitivo | Raro | Root/abas/import/export | OK no root. |
| `documentData` | Objeto grande OT | Por acao/commit buffered | Documento OT, permissoes, testes, review, preview | Candidato futuro a reducer/store por dominio se Profiler apontar root quente. |
| `teaData` | Objeto grande TEA | Por acao/commit buffered | TEA document/activities/review/preview | Maior candidato a colocation parcial. |
| `otPreviewDocumentData`, `teaPreviewDocumentData` | Snapshot grande | Manual ao atualizar preview | Preview | OK; evita rebuild continuo do preview. |
| `expandedTests` | Mapa medio/grande | Toggle | Testes OT | Pode descer para painel de testes. |
| `collapsedMacros`, `collapsedPermissionBlocks` | Mapas medios | Toggle | Testes OT | Pode descer para painel de testes. |
| `collapsedTeaActivities`, `collapsedTeaSubActivities`, `collapsedTeaComposers`, `collapsedTeaContentBlocks` | Mapas medios/grandes | Toggle | TEA activities | Melhor candidato a mover para `TeaWorkspace`. |
| `activeTab`, `teaActiveTab`, `testBlockFilter` | Primitivos | Clique | Navegacao/filtros | OK; `testBlockFilter` pode descer se painel OT for extraido. |
| Flags de loading/import/export/confirm/copy | Primitivos | Raro | Modais/topbar | OK no root. |
| `importPreview`, `pendingConfirmation` | Objetos pequenos/medios | Raro | Modais | OK no root. |
| `draftStatus`, `permissionBulkText` | Primitivos/string | Raro depois do Round 1 | Topbar/permissoes | OK. |

## Effects auditados em A4

Ja estavam corretos:

- Autosave do rascunho: depende de dados atuais e callbacks estabilizados; limpa timeout/idle callback.
- `beforeunload`: registra listener com cleanup e callbacks estaveis.
- Sync de `permissionBulkTextRef`: dependencia primitiva.
- `DocumentOutline`: scroll apenas quando `activeTargetId` muda.
- `LazyCollapse`: timeout limpo no cleanup.
- `useBufferedText`: mantem `onCommit` em ref, limpa timeout e registra commit buffered com cleanup.

Corrigido:

- `useActiveOutlineTargetId`: o effect agora deriva `scopedTargetIds` de `targetSignature`, evitando dependencia ausente em `targetIds` sem recriar listeners por identidade de array.

Documentado:

- Hydration inicial de imagens: `[]` intencional, executa apenas na montagem.
