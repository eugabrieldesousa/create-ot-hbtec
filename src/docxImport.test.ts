import { describe, expect, it } from "vitest";
import { parseOtHtml } from "./docxImport";

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
    expect(macro.microPermissions.map((micro) => micro.code)).toEqual(["AT", "SC"]);

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
      selectedPermissions: 2,
      tests: 1,
      images: 2,
    });
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
});
