import { describe, expect, it } from "vitest";
import { parseOtHtml, parseTeaHtml } from "./docxImport";

const pngDataUrl =
  "data:image/png;base64,iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mP8z8BQDwAFgwJ/lSZ8nAAAAABJRU5ErkJggg==";

describe("parseOtHtml", () => {
  it("imports the OT model with metadata, permissions, tests, checks and images", () => {
    const result = parseOtHtml(
      `
        <h2>OBSERVABILIDADE DE TESTES (PERMISSÕES + NAVEGAÇÃO)</h2>
        <p><strong>Objetivo:</strong> Garantir que a tela esteja funcionando.</p>
        <table>
          <tr><td>Tela:</td><td>CADASTRO DE INDICADORES</td></tr>
          <tr><td>Responsável pelo teste:</td><td>GABRIEL</td></tr>
          <tr><td>Data:</td><td>29/04/2026</td></tr>
          <tr><td>Ambiente:</td><td>LOCAL + DESENVOLVIMENTO</td></tr>
          <tr><td>Elaboradora por:</td><td>GABRIEL</td></tr>
        </table>
        <h2>Passo a Passo para Acessar a Tela</h2>
        <ol>
          <li>Acessar o sistema</li>
          <li>Selecionar uma operadora</li>
          <li>Clicar em: Qualificacao &gt; plano de visitas</li>
        </ol>
        <h2>Tipos de Permissão para Testar</h2>
        <p><strong>Micro-permissões:</strong></p>
        <ul>
          <li>AT (Atualização)</li>
          <li>SC (Somente Consulta)</li>
        </ul>
        <p><strong>Macro-permissões (Tipo de Usuário):</strong></p>
        <p>AO</p>
        <h2>TESTES</h2>
        <table>
          <tr><td>MACRO-PERMISSÃO</td><td>Tipo de usuário: AO (Administrador Geral)</td></tr>
          <tr><td>MICRO-PERMISSÃO</td><td>Tipo de permissão: AT (Atualização)</td></tr>
        </table>
        <table>
          <tr><td colspan="2">1 - FILTRO</td></tr>
          <tr><td>( X )</td><td>Funcionou legado e novo estão com o mesmo comportamento</td></tr>
          <tr><td>(   )</td><td>Possível problema/Comportamento estranho/Ambiguidade/incoerência</td></tr>
          <tr><td>( X )</td><td>Problema em AMBOS</td></tr>
          <tr><td>(   )</td><td>Problema só no NOVO</td></tr>
          <tr><td>( X )</td><td>Relatório de Erros</td></tr>
        </table>
        <table>
          <tr><td>Observações:</td><td>Tirar observações e ajustar realizada em</td></tr>
        </table>
        <p><strong>Legado:</strong></p>
        <p><strong>com</strong><img src="${pngDataUrl}" /></p>
        <p><strong>Novo:</strong></p>
        <p><img src="${pngDataUrl}" /></p>
      `,
      { sourceName: "OT - Plano de Visitas.docx" },
    );

    expect(result.document.metadata).toMatchObject({
      screen: "CADASTRO DE INDICADORES",
      responsible: "GABRIEL",
      date: "2026-04-29",
      environment: "LOCAL + DESENVOLVIMENTO",
      author: "GABRIEL",
    });
    expect(result.document.objective).toBe("Garantir que a tela esteja funcionando.");
    expect(result.document.accessSteps.map((step) => step.text)).toEqual([
      "Acessar o sistema",
      "Selecionar uma operadora",
      "Clicar em: Qualificacao > plano de visitas",
    ]);

    const macro = result.document.permissionGroups[0];
    expect(macro.code).toBe("AO");
    expect(macro.label).toBe("Administrador Geral");
    expect(macro.microPermissions.map((micro) => micro.code)).toEqual(["AT"]);

    const activeMicro = macro.microPermissions[0];
    const block = result.document.permissionBlocks[`${macro.id}:${activeMicro.id}`];
    const test = block.tests[0];

    expect(test.title).toBe("FILTRO");
    expect(test.result.checks).toMatchObject({
      sameBehavior: true,
      possibleIssue: false,
      bothIssue: true,
      newIssue: false,
      errorReport: true,
    });
    expect(test.result.observations).toBe("Tirar observações e ajustar realizada em");
    expect(test.result.legacyImages).toHaveLength(1);
    expect(test.result.legacyImages[0].label).toBe("com");
    expect(test.result.newImages).toHaveLength(1);
    expect(result.summary).toMatchObject({
      accessSteps: 3,
      permissionGroups: 1,
      selectedPermissions: 1,
      tests: 1,
      images: 2,
    });
  });

  it("imports current quick status labels without breaking legacy check keys", () => {
    const result = parseOtHtml(
      `
        <h2>TESTES</h2>
        <table>
          <tr><td>MACRO-PERMISSÃO</td><td>Tipo de usuário: AO (Administrador Geral)</td></tr>
          <tr><td>MICRO-PERMISSÃO</td><td>Tipo de permissão: AT (Atualização)</td></tr>
        </table>
        <table>
          <tr><td colspan="2">1 - STATUS ATUAL</td></tr>
          <tr><td>(   )</td><td>Funcionou legado e novo estão com o mesmo comportamento</td></tr>
          <tr><td>( X )</td><td>Possível problema de lógica ou regra de negócio</td></tr>
          <tr><td>( X )</td><td>Erro no legado</td></tr>
          <tr><td>( X )</td><td>Erro no novo</td></tr>
          <tr><td>( X )</td><td>Relatório de Erros</td></tr>
        </table>
      `,
      { sourceName: "OT - Status atual.docx" },
    );

    const macro = result.document.permissionGroups[0];
    const micro = macro.microPermissions[0];
    const test = result.document.permissionBlocks[`${macro.id}:${micro.id}`].tests[0];

    expect(test.result.checks).toMatchObject({
      sameBehavior: false,
      possibleIssue: true,
      bothIssue: true,
      newIssue: true,
      errorReport: true,
    });
  });

  it("does not add summary-only micro permissions to a tested OT macro", () => {
    const result = parseOtHtml(
      `
        <h2>Tipos de Permissão para Testar</h2>
        <p><strong>Micro-permissões:</strong></p>
        <ul>
          <li>AT (Atualização)</li>
          <li>SC (Somente Consulta)</li>
          <li>SA (Sem Acesso)</li>
        </ul>
        <p><strong>Macro-permissões (Tipo de Usuário):</strong></p>
        <p>AO</p>
        <h2>TESTES</h2>
        <table>
          <tr><td>MACRO-PERMISSÃO</td><td>Tipo de usuário: AO (Administrador Geral)</td></tr>
          <tr><td>MICRO-PERMISSÃO</td><td>Tipo de permissão: AT (Atualização)</td></tr>
        </table>
        <table>
          <tr><td colspan="2">1 - CADASTRO</td></tr>
          <tr><td>( X )</td><td>Funcionou legado e novo estão com o mesmo comportamento</td></tr>
        </table>
      `,
      { sourceName: "OT - Permissoes.docx" },
    );

    const macro = result.document.permissionGroups[0];

    expect(macro.code).toBe("AO");
    expect(macro.microPermissions.map((micro) => micro.code)).toEqual(["AT"]);
    expect(result.summary.selectedPermissions).toBe(1);
  });

  it("handles a partial DOCX without images or tests", () => {
    const result = parseOtHtml(
      `
        <p><strong>Objetivo:</strong> Conferir cadastro.</p>
        <table>
          <tr><td>Tela:</td><td>Cadastro simples</td></tr>
          <tr><td>Data:</td><td>2026-06-06</td></tr>
        </table>
        <h2>Passo a Passo para Acessar a Tela</h2>
        <p>1. Entrar no menu 2. Abrir cadastro</p>
        <h2>Tipos de Permissão para Testar</h2>
        <table>
          <tr><td>Macro-permissão</td><td>PR (Prestador)</td></tr>
          <tr><td>Micro-permissão</td><td>SA (Sem Acesso)</td></tr>
        </table>
      `,
      { sourceName: "OT - Cadastro simples.docx" },
    );

    expect(result.document.metadata.screen).toBe("Cadastro simples");
    expect(result.document.accessSteps.map((step) => step.text)).toEqual([
      "Entrar no menu",
      "Abrir cadastro",
    ]);
    expect(result.document.permissionGroups[0].code).toBe("PR");
    expect(result.document.permissionGroups[0].microPermissions[0].code).toBe("SA");
    expect(result.summary.tests).toBe(0);
    expect(result.summary.images).toBe(0);
    expect(result.warnings).toContain("Nenhum teste reconhecido.");
  });

  it("imports correction details and before/after evidence from OT DOCX", () => {
    const result = parseOtHtml(
      `
        <h2>TESTES</h2>
        <table>
          <tr><td>MACRO-PERMISSÃO</td><td>Tipo de usuário: AO (Administrador Geral)</td></tr>
          <tr><td>MICRO-PERMISSÃO</td><td>Tipo de permissão: AT (Atualização)</td></tr>
        </table>
        <table>
          <tr><td colspan="2">1 - QUESTIONARIO</td></tr>
          <tr><td>(   )</td><td>Funcionou legado e novo estão com o mesmo comportamento</td></tr>
          <tr><td>(   )</td><td>Possível problema/Comportamento estranho/Ambiguidade/incoerência</td></tr>
          <tr><td>(   )</td><td>Problema em AMBOS</td></tr>
          <tr><td>( X )</td><td>Problema só no NOVO</td></tr>
          <tr><td>(   )</td><td>Relatório de Erros</td></tr>
        </table>
        <p><strong>Correcao:</strong></p>
        <table>
          <tr><td>Corrigido</td><td>Sim</td></tr>
          <tr><td>Hotfix</td><td>hotfix 1.2.2</td></tr>
          <tr><td>Corrigido por</td><td>Gabriel Sousa</td></tr>
          <tr><td>Nuvem</td><td>Ate homolog</td></tr>
        </table>
        <p><strong>Antes (com erro):</strong></p>
        <p>Falha antes<img src="${pngDataUrl}" /></p>
        <p><strong>Depois (corrigido):</strong></p>
        <p>Ok depois<img src="${pngDataUrl}" /></p>
      `,
      { sourceName: "OT - Questionario.docx" },
    );

    const macro = result.document.permissionGroups[0];
    const micro = macro.microPermissions[0];
    const correction = result.document.permissionBlocks[`${macro.id}:${micro.id}`].tests[0]
      .correction;

    expect(correction).toMatchObject({
      corrected: true,
      hotfixTag: "hotfix 1.2.2",
      correctedBy: "Gabriel Sousa",
      cloudStage: "homolog",
    });
    expect(correction?.beforeImages).toHaveLength(1);
    expect(correction?.beforeImages[0].label).toBe("Falha antes");
    expect(correction?.afterImages).toHaveLength(1);
    expect(correction?.afterImages[0].label).toBe("Ok depois");
    expect(result.summary.images).toBe(2);
  });

  it("imports current OT error cards with error prints and correction evidence", () => {
    const result = parseOtHtml(
      `
        <h2>TESTES</h2>
        <table>
          <tr><td>MACRO-PERMISSÃO</td><td>Tipo de usuário: AO (Administrador Geral)</td></tr>
          <tr><td>MICRO-PERMISSÃO</td><td>Tipo de permissão: AT (Atualização)</td></tr>
        </table>
        <table>
          <tr><td colspan="2">1 - QUESTIONARIO</td></tr>
          <tr><td>(   )</td><td>Funcionou legado e novo estão com o mesmo comportamento</td></tr>
          <tr><td>(   )</td><td>Possível problema de lógica ou regra de negócio</td></tr>
          <tr><td>(   )</td><td>Erro no legado</td></tr>
          <tr><td>( X )</td><td>Erro no novo</td></tr>
          <tr><td>( X )</td><td>Relatório de Erros</td></tr>
        </table>
        <p><strong>Erros encontrados:</strong></p>
        <p><strong>Erro 1 - Novo</strong></p>
        <table>
          <tr><td>Origem</td><td>Novo</td></tr>
          <tr><td>Observacao</td><td>Falha ao salvar no novo.</td></tr>
        </table>
        <p><strong>Prints do erro:</strong></p>
        <p>Falha visivel<img src="${pngDataUrl}" /></p>
        <p><strong>Correcao:</strong></p>
        <table>
          <tr><td>Corrigido</td><td>Sim</td></tr>
          <tr><td>Hotfix</td><td>hotfix 2.0.0</td></tr>
          <tr><td>Corrigido por</td><td>Gabriel Sousa</td></tr>
          <tr><td>Nuvem</td><td>Ate homolog</td></tr>
        </table>
        <p><strong>Antes (com erro):</strong></p>
        <p>Antes da correcao<img src="${pngDataUrl}" /></p>
        <p><strong>Depois (corrigido):</strong></p>
        <p>Depois da correcao<img src="${pngDataUrl}" /></p>
      `,
      { sourceName: "OT - Questionario.docx" },
    );

    const macro = result.document.permissionGroups[0];
    const micro = macro.microPermissions[0];
    const test = result.document.permissionBlocks[`${macro.id}:${micro.id}`].tests[0];
    const error = test.result.errors[0];

    expect(test.result.legacyImages).toHaveLength(0);
    expect(test.result.newImages).toHaveLength(0);
    expect(test.result.errors).toHaveLength(1);
    expect(error).toMatchObject({
      origin: "new",
      observation: "Falha ao salvar no novo.",
    });
    expect(error.images).toHaveLength(1);
    expect(error.images[0].label).toBe("Falha visivel");
    expect(error.correction).toMatchObject({
      corrected: true,
      hotfixTag: "hotfix 2.0.0",
      correctedBy: "Gabriel Sousa",
      cloudStage: "homolog",
    });
    expect(error.correction.beforeImages[0].label).toBe("Antes da correcao");
    expect(error.correction.afterImages[0].label).toBe("Depois da correcao");
    expect(result.summary.images).toBe(3);
  });
});

describe("parseTeaHtml", () => {
  it("imports the TEA model with metadata, sections, content blocks and images", () => {
    const result = parseTeaHtml(
      `
        <p><strong>Termo de Entrega de Atividade (TEA)</strong></p>
        <table>
          <tr><td>Ordem de Serviço:</td><td>OS2171 - Documentos Vencidos</td></tr>
          <tr><td>Fase/Etapa</td><td>Etapa 5</td></tr>
          <tr><td>Chamado:</td><td>Chamado 202504000396</td></tr>
          <tr><td>Assunto:</td><td>Telas - Novo Layout</td></tr>
          <tr><td>Data:</td><td>28/04/2026</td></tr>
          <tr><td>Elaborado por:</td><td>Gabriel Sousa</td></tr>
        </table>
        <h1>1. VISÃO GERAL</h1>
        <p>Visão geral importada.</p>
        <h1>2. ATIVIDADES REALIZADAS</h1>
        <p>Texto inicial das atividades.</p>
        <p><img src="${pngDataUrl}" /></p>
        <h2>2.1 - Botão Editar:</h2>
        <p>Texto da atividade.</p>
        <ul>
          <li>Item da atividade</li>
          <li>Outro item</li>
        </ul>
        <p><img src="${pngDataUrl}" /></p>
        <p>2.1.1 - Modal de edição:</p>
        <p>Texto da subatividade.</p>
        <ul>
          <li>Item da subatividade</li>
        </ul>
      `,
      { sourceName: "TEA - Documentos Vencidos.docx" },
    );

    expect(result.kind).toBe("tea");
    expect(result.document.metadata).toMatchObject({
      serviceOrder: "OS2171 - Documentos Vencidos",
      phase: "Etapa 5",
      ticket: "Chamado 202504000396",
      subject: "Telas - Novo Layout",
      date: "2026-04-28",
      author: "Gabriel Sousa",
    });
    expect(result.document.overview).toBe("Visão geral importada.");
    expect(result.document.activityIntro).toBe("Texto inicial das atividades.");
    expect(result.document.activityImages).toHaveLength(1);

    const activity = result.document.activities[0];
    expect(activity.title).toBe("Botão Editar");
    expect(activity.blocks.map((block) => block.type)).toEqual(["text", "list", "images"]);
    expect(activity.blocks[0]).toMatchObject({
      type: "text",
      text: "Texto da atividade.",
    });
    expect(activity.blocks[1]).toMatchObject({
      type: "list",
      items: [{ text: "Item da atividade" }, { text: "Outro item" }],
    });

    const subActivity = activity.subActivities[0];
    expect(subActivity.title).toBe("Modal de edição");
    expect(subActivity.blocks.map((block) => block.type)).toEqual(["text", "list"]);
    expect(result.summary).toMatchObject({
      subject: "Telas - Novo Layout",
      activities: 1,
      subActivities: 1,
      blocks: 5,
      images: 2,
    });
    expect(result.warnings).toEqual([]);
  });

  it("uses the source name and reports warnings when TEA sections are incomplete", () => {
    const result = parseTeaHtml(
      `
        <h1>1. VISÃO GERAL</h1>
        <p>Visão geral importada.</p>
      `,
      { sourceName: "TEA - Documento Parcial.docx" },
    );

    expect(result.document.metadata.subject).toBe("Documento Parcial");
    expect(result.warnings).toContain("Assunto nao encontrado no DOCX; usei o nome do arquivo.");
    expect(result.warnings).toContain("Nenhuma atividade reconhecida.");
  });

  it("preserves bold TEA text as double-asterisk markup without affecting headings or metadata", () => {
    const result = parseTeaHtml(
      `
        <table>
          <tr><td>Ordem de servico:</td><td><strong>OS2171</strong></td></tr>
          <tr><td>Assunto:</td><td><strong>Telas - Novo Layout</strong></td></tr>
          <tr><td>Elaborado por:</td><td><strong>Gabriel Sousa</strong></td></tr>
        </table>
        <h1><strong>1. VISAO GERAL</strong></h1>
        <p>Visao <strong>geral</strong> importada.</p>
        <h1><strong>2. ATIVIDADES REALIZADAS</strong></h1>
        <p>Texto inicial das <strong>atividades</strong>.</p>
        <h2><strong>2.1 - Botao Editar:</strong></h2>
        <p>Texto da <strong>atividade</strong>.</p>
        <ul>
          <li>Item da <strong>atividade</strong></li>
        </ul>
      `,
      { sourceName: "TEA - Documentos Vencidos.docx" },
    );

    expect(result.document.metadata.serviceOrder).toBe("OS2171");
    expect(result.document.metadata.subject).toBe("Telas - Novo Layout");
    expect(result.document.overview).toBe("Visao **geral** importada.");
    expect(result.document.activityIntro).toBe("Texto inicial das **atividades**.");

    const activity = result.document.activities[0];
    expect(activity.title).toBe("Botao Editar");
    expect(activity.blocks[0]).toMatchObject({
      type: "text",
      text: "Texto da **atividade**.",
    });
    expect(activity.blocks[1]).toMatchObject({
      type: "list",
      items: [{ text: "Item da **atividade**" }],
    });
  });

  it("imports externally added TEA activity blocks with unique editable identities", () => {
    const result = parseTeaHtml(
      `
        <h1>1. VISAO GERAL</h1>
        <p>Visao geral importada.</p>
        <h1>2. ATIVIDADES REALIZADAS</h1>
        <p>Texto inicial.</p>
        <h2>2.1 - Atividade repetida:</h2>
        <p>Primeiro paragrafo externo.</p>
        <p>Segundo paragrafo externo.</p>
        <ul>
          <li>Item externo</li>
        </ul>
        <h2>2.2 - Atividade repetida:</h2>
        <p>Outro bloco externo.</p>
      `,
      { sourceName: "TEA - Atividade Externa.docx" },
    );

    const [firstActivity, secondActivity] = result.document.activities;
    const firstBlockIds = firstActivity.blocks.map((block) => block.id);
    const activityIds = result.document.activities.map((activity) => activity.id);

    expect(firstActivity.blocks.map((block) => block.type)).toEqual(["text", "text", "list"]);
    expect(firstActivity.blocks[0]).toMatchObject({
      type: "text",
      text: "Primeiro paragrafo externo.",
    });
    expect(firstActivity.blocks[1]).toMatchObject({
      type: "text",
      text: "Segundo paragrafo externo.",
    });
    expect(new Set(firstBlockIds).size).toBe(firstBlockIds.length);
    expect(new Set(activityIds).size).toBe(activityIds.length);
    expect(secondActivity.title).toBe("Atividade repetida");
  });

  it("ignores Mammoth's built-in Title style warning", () => {
    const result = parseTeaHtml(
      `
        <h1>1. VISÃO GERAL</h1>
        <p>Visão geral importada.</p>
      `,
      {
        sourceName: "TEA - Documento Parcial.docx",
        mammothMessages: [
          {
            message: "Unrecognised paragraph style: 'Title' (Style ID: Title)",
          },
        ],
      },
    );

    expect(result.warnings).not.toContain(
      "Conversao DOCX: Unrecognised paragraph style: 'Title' (Style ID: Title)",
    );
  });
});
