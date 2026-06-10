# Analise de Performance - React Application

Data: 2026-06-10

## Resumo executivo

O maior custo de performance estava concentrado em `src/App.tsx`: uma arvore React grande, muitos paineis de collapse aninhados, filtros de pendencias dentro de loops de render e algumas props recriadas a cada render. Os top 5 abaixo foram corrigidos com mudancas cirurgicas. A logica de negocio foi preservada.

Validacao executada:

```bash
npm run build
npm test
```

Resultado: build OK, 6 arquivos de teste OK, 55 testes OK.

## Lista priorizada

### 1. Props e filtros de pendencias recriados em listas grandes

**Arquivo:** `src/App.tsx`

**Problema:** TEA e OT filtravam `reviewSummary.issues` dentro de loops (`activities`, `subActivities`, `blocks`, `tests`). Isso criava arrays novos em cada render e invalidava `React.memo`.

**Impacto:** alto

**Causa raiz:** `Array.filter` dentro do JSX gera custo O(lista x pendencias) e cria referencias novas. Mesmo quando os dados do card nao mudavam, filhos memoizados recebiam props novas.

**Solucao aplicada:** foram adicionados indices memoizados:

```tsx
const reviewIssueIndex = useMemo(
  () => buildReviewIssueIndex(reviewSummary.issues),
  [reviewSummary.issues],
);

const teaReviewIssueIndex = useMemo(
  () => buildTeaReviewIssueIndex(teaReviewSummary.issues),
  [teaReviewSummary.issues],
);
```

**Ganho esperado:** menos re-renders em atividades, subtitulos, blocos e testes; melhor resposta em expand/collapse e revisao.

### 2. `React.memo` inefetivo nos cards TEA

**Arquivo:** `src/App.tsx`

**Problema:** `TeaActivityEditor`, `TeaSubActivityEditor`, `TeaContentComposer` e `TeaContentBlockEditor` eram memoizados, mas recebiam callbacks inline e mapas globais de collapse. Qualquer toggle de um item podia re-renderizar irmaos.

**Impacto:** alto

**Causa raiz:** shallow compare do `memo` falha quando funcoes/objetos mudam por identidade, mesmo sem mudanca relevante para o card.

**Solucao aplicada:** comparadores especificos por item:

```tsx
}, areTeaActivityEditorPropsEqual);
}, areTeaSubActivityEditorPropsEqual);
}, areTeaContentComposerPropsEqual);
}, areTeaContentBlockEditorPropsEqual);
```

**Ganho esperado:** ao recolher uma atividade/bloco, os irmaos que nao mudaram deixam de renderizar. Isso reduz jank em documentos TEA maiores.

**Possivel efeito colateral:** os comparadores assumem que callbacks inline mantem a mesma semantica quando o id do item nao muda. Esse padrao esta valido no codigo atual.

### 3. Subarvores TEA continuavam montadas apos collapse

**Arquivo:** `src/App.tsx`, `src/styles.css`

**Problema:** paineis TEA fechados mantinham inputs, uploaders e blocos filhos montados dentro do `Collapse`.

**Impacto:** alto

**Causa raiz:** `Collapse` anima altura e precisa medir layout. Com uma subarvore grande montada, renders e medicoes ficam caros mesmo quando o painel parece fechado.

**Solucao aplicada:** novo `LazyCollapse`, que mantem o conteudo durante a animacao e desmonta apos a transicao:

```tsx
<LazyCollapse in={isExpanded}>
  <div id={panelId} className="collapseBody">
    ...
  </div>
</LazyCollapse>
```

CSS aplicado:

```css
.collapseBody {
  contain: layout paint;
  transform: translateZ(0);
  will-change: transform, opacity;
}
```

**Ganho esperado:** menos memoria e menos trabalho de reconciliacao quando paineis estao recolhidos; collapse mais fluido.

**Possivel efeito colateral:** estado local interno de um painel fechado e desmontado e descartado. Campos baseados em `useBufferedText` ja fazem commit no cleanup.

### 4. Campo "Permissoes em lote" re-renderizava o App a cada tecla

**Arquivo:** `src/App.tsx`

**Problema:** a textarea de permissoes em lote era controlada por estado no `App`.

**Impacto:** alto para digitacao nesse campo

**Causa raiz:** `setPermissionBulkText` no `onChange` disparava render da arvore principal em cada caractere.

**Solucao aplicada:** o campo virou uncontrolled com `defaultValue`; o texto corrente fica em `permissionBulkTextRef` e so e commitado no blur/apply/load.

```tsx
<Textarea
  key={permissionBulkText}
  defaultValue={permissionBulkText}
  onBlur={commitPermissionBulkDraft}
  onChange={(event) => updatePermissionBulkDraft(event.currentTarget.value)}
/>
```

**Ganho esperado:** digitacao no editor em lote nao re-renderiza toda a aplicacao por tecla.

### 5. Efeito de indice reiniciava listeners por identidade de array

**Arquivo:** `src/App.tsx`

**Problema:** `useActiveOutlineTargetId` dependia de `targetIds` e de `targetSignature`. Mesmo quando a assinatura era igual, uma nova referencia de array podia recriar listeners de scroll/resize.

**Impacto:** medio

**Causa raiz:** dependencia redundante em array derivado.

**Solucao aplicada:** o efeito passa a depender de `scopeKey` e `targetSignature`.

**Ganho esperado:** menos churn de listener em mudancas de revisao/outline.

### 6. Animacoes de icone sem promocao de camada

**Arquivo:** `src/styles.css`

**Problema:** o icone do toggle rotacionava sem dica de composicao.

**Impacto:** medio

**Causa raiz:** animacao de `transform` sem `will-change` pode disputar trabalho com layout em telas carregadas.

**Solucao aplicada:**

```css
.testToggleIcon {
  transform: translateZ(0);
  transition: transform 160ms ease;
  will-change: transform;
}
```

**Ganho esperado:** rotacao do chevron mais estavel durante expand/collapse.

### 7. Bundle principal muito grande

**Arquivo:** `src/App.tsx`, `src/docxExport.ts`, `src/docxImport.ts`, `src/docxPreview.tsx`

**Problema:** build gerou chunk principal com cerca de 1.41 MB minificado e 399 KB gzip.

**Impacto:** medio/alto no carregamento inicial

**Causa raiz:** exportacao DOCX, importacao DOCX, preview e UI principal entram no mesmo grafo inicial.

**Solucao recomendada:** aplicar `React.lazy`/dynamic import para preview e importar `docxExport`/`docxImport` sob demanda nos handlers de export/import.

**Ganho esperado:** menor TTI inicial; custo de DOCX carregado apenas quando usuario importa/exporta/abre preview.

### 8. Listas longas ainda nao usam virtualizacao

**Arquivo:** `src/App.tsx`

**Problema:** atividades, blocos, testes e imagens renderizam todos os itens quando o usuario expande tudo.

**Impacto:** medio; alto em documentos muito grandes

**Causa raiz:** renderizacao direta com `.map`.

**Solucao recomendada:** se documentos reais passarem de ~50 cards expandidos, usar TanStack Virtual ou `react-window` nas listas de atividades/testes/imagens.

**Ganho esperado:** custo proporcional ao viewport, nao ao total de itens.

### 9. `PermissionGroupEditor` ainda nao e memoizado

**Arquivo:** `src/App.tsx`

**Problema:** macros e micros de permissao ainda recebem callbacks inline.

**Impacto:** medio em listas grandes de permissao

**Causa raiz:** componente sem `memo` e sem comparador por macro.

**Solucao recomendada:** extrair props tipadas e memoizar `PermissionGroupEditor`, ou criar handlers por id em um componente filho.

**Ganho esperado:** menos re-render ao editar uma macro/micro isolada.

### 10. Preview DOCX usa chaves parcialmente baseadas em indice

**Arquivo:** `src/docxPreview.tsx`

**Problema:** algumas keys usam `index`, `rowIndex`, `cellIndex` e `paragraphIndex`.

**Impacto:** baixo/medio, restrito a aba de preview

**Causa raiz:** quando blocos sao inseridos/removidos, React pode reconciliar mais nos trechos seguintes.

**Solucao recomendada:** carregar ids estaveis no model para listas/tabelas/runs quando o preview ficar maior.

**Ganho esperado:** menos reconciliacao em previews grandes.

## Checklist progressivo

- Medir antes/depois com React DevTools Profiler nos fluxos TEA: abrir/recolher atividade, abrir/recolher bloco, digitar em bloco de texto.
- Fazer code splitting de `docxExport`, `docxImport` e preview.
- Memoizar `PermissionGroupEditor` se macros/micros crescerem.
- Avaliar virtualizacao quando houver documentos reais grandes.
- Evitar novos `.filter/.map/.sort` diretamente em JSX quando o resultado for prop de filho memoizado.
- Manter inputs longos com estado local, `useBufferedText` ou uncontrolled + commit explicito.
- Evitar `Collapse` mantendo subarvore pesada montada indefinidamente.

## Profiling recomendado para este projeto

### React DevTools Profiler

1. Rode `npm run dev`.
2. Abra a app no navegador.
3. No React DevTools, aba Profiler, habilite "Record why each component rendered".
4. Grave estes fluxos:
   - TEA > Atividades > Expandir todos > Recolher todos.
   - Editar texto em um bloco TEA por 5 segundos.
   - OT > Permissoes > digitar no campo "Permissoes em lote".
5. Verifique se apenas o card editado aparece como render quente.

### Chrome Performance tab

1. Abra DevTools > Performance.
2. Marque Screenshots e Web Vitals.
3. Grave expand/collapse em um documento grande.
4. Procure por long tasks acima de 50 ms, Layout e Recalculate Style durante `Collapse`.
5. Compare antes/depois pelo tempo de scripting/layout no frame do toggle.

### why-did-you-render

Use apenas em desenvolvimento local. Sugestao:

```tsx
// src/main.tsx, somente em dev
if (import.meta.env.DEV) {
  // instalar @welldone-software/why-did-you-render antes
}
```

Alvos iniciais:

- `TeaActivityEditor`
- `TeaSubActivityEditor`
- `TeaContentComposer`
- `TeaContentBlockEditor`
- `PermissionBlockGroup`
- `PermissionBlockEditor`

### Bundle analyzer

Use `vite-bundle-visualizer` ou `rollup-plugin-visualizer` para confirmar o peso de `docx`, `mammoth`, Mantine e preview. O build atual ja alerta chunk acima de 500 KB.
