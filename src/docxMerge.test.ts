import { describe, expect, it } from "vitest";
import {
  applyOtDocxMerge,
  applyTeaDocxMerge,
  findMatchingOtMergeTarget,
} from "./docxMerge";
import type { OtDocument, TeaDocument } from "./types";

describe("applyTeaDocxMerge", () => {
  it("inserts selected TEA activities at the configured position with new ids", () => {
    const current = createTeaDocument("Atual", [
      createTeaActivity("current-a", "Atual A"),
      createTeaActivity("current-b", "Atual B"),
    ]);
    const source = createTeaDocument("Importado", [
      createTeaActivity("source-a", "Importada A", [
        createTeaTextBlock("source-a-text", "Texto importado"),
      ]),
    ]);

    const result = applyTeaDocxMerge(
      current,
      source,
      {
        activityIds: ["source-a"],
        subActivityIds: [],
        activityPosition: { mode: "before", targetId: "current-b" },
        subActivityTargetActivityId: null,
        subActivityPosition: { mode: "end" },
      },
      idFactory("new"),
    );

    expect(result.document.metadata.subject).toBe("Atual");
    expect(result.document.activities.map((activity) => activity.title)).toEqual([
      "Atual A",
      "Importada A",
      "Atual B",
    ]);
    expect(result.document.activities[1].id).toBe("new-1");
    expect(result.document.activities[1].blocks[0].id).toBe("new-2");
    expect(source.activities[0].id).toBe("source-a");
  });

  it("inserts selected loose TEA subactivities into an existing activity", () => {
    const current = createTeaDocument("Atual", [
      createTeaActivity("current-a", "Atual A", [], [
        createTeaSubActivity("current-sub-a", "Sub atual A"),
        createTeaSubActivity("current-sub-b", "Sub atual B"),
      ]),
    ]);
    const source = createTeaDocument("Importado", [
      createTeaActivity("source-a", "Importada A", [], [
        createTeaSubActivity("source-sub-a", "Sub importado", [
          createTeaListBlock("source-list", "source-item", "Item importado"),
        ]),
      ]),
    ]);

    const result = applyTeaDocxMerge(
      current,
      source,
      {
        activityIds: [],
        subActivityIds: ["source-sub-a"],
        activityPosition: { mode: "end" },
        subActivityTargetActivityId: "current-a",
        subActivityPosition: { mode: "after", targetId: "current-sub-a" },
      },
      idFactory("sub"),
    );

    const subActivities = result.document.activities[0].subActivities;
    expect(result.document.metadata.subject).toBe("Atual");
    expect(subActivities.map((subActivity) => subActivity.title)).toEqual([
      "Sub atual A",
      "Sub importado",
      "Sub atual B",
    ]);
    expect(subActivities[1].id).toBe("sub-1");
    expect(subActivities[1].blocks[0].id).toBe("sub-2");
    expect(subActivities[1].blocks[0]).toMatchObject({
      type: "list",
      items: [{ id: "sub-3", text: "Item importado" }],
    });
  });
});

describe("applyOtDocxMerge", () => {
  it("inserts selected OT tests into an existing permission block", () => {
    const current = createOtDocument("Tela atual", [
      createPermissionGroup("macro-a", "AO", "Administrador", [
        createPermission("micro-a", "AT", "Atualizacao"),
      ]),
    ], {
      "macro-a:micro-a": {
        tests: [createTest("current-test", "Teste atual")],
      },
    });
    const source = createOtDocument("Tela importada", [
      createPermissionGroup("source-macro", "AO", "Administrador", [
        createPermission("source-micro", "AT", "Atualizacao"),
      ]),
    ], {
      "source-macro:source-micro": {
        tests: [
          createTest("source-test-a", "Teste importado A", "source-image-a"),
          createTest("source-test-b", "Teste importado B"),
        ],
      },
    });

    const result = applyOtDocxMerge(
      current,
      source,
      {
        groups: [
          {
            sourceMacroId: "source-macro",
            sourceMicroId: "source-micro",
            testIds: ["source-test-a"],
            target: { kind: "existing", macroId: "macro-a", microId: "micro-a" },
            position: { mode: "before", targetId: "current-test" },
          },
        ],
      },
      idFactory("ot"),
    );

    const tests = result.document.permissionBlocks["macro-a:micro-a"].tests;
    expect(result.document.metadata.screen).toBe("Tela atual");
    expect(tests.map((test) => test.title)).toEqual(["Teste importado A", "Teste atual"]);
    expect(tests[0].id).toBe("ot-1");
    expect(tests[0].result.legacyImages[0].id).toBe("ot-2");
    expect(current.permissionBlocks["macro-a:micro-a"].tests).toHaveLength(1);
  });

  it("creates OT permissions for selected tests when no target exists", () => {
    const current = createOtDocument("Tela atual", [], {});
    const source = createOtDocument("Tela importada", [
      createPermissionGroup("source-macro", "NEW", "Nova macro", [
        createPermission("source-micro", "IN", "Inclusao"),
      ]),
    ], {
      "source-macro:source-micro": {
        tests: [createTest("source-test", "Teste novo", "source-image")],
      },
    });

    const result = applyOtDocxMerge(
      current,
      source,
      {
        groups: [
          {
            sourceMacroId: "source-macro",
            sourceMicroId: "source-micro",
            testIds: ["source-test"],
            target: { kind: "new" },
            position: { mode: "end" },
          },
        ],
      },
      idFactory("created"),
    );

    expect(result.document.permissionGroups).toEqual([
      {
        id: "created-1",
        code: "NEW",
        label: "Nova macro",
        selected: true,
        microPermissions: [
          {
            id: "created-2",
            code: "IN",
            label: "Inclusao",
            selected: true,
          },
        ],
      },
    ]);
    expect(result.document.permissionBlocks["created-1:created-2"].tests[0]).toMatchObject({
      id: "created-3",
      title: "Teste novo",
      result: {
        legacyImages: [{ id: "created-4" }],
      },
    });
  });

  it("finds matching OT merge targets by permission code", () => {
    const current = createOtDocument("Tela atual", [
      createPermissionGroup("macro-a", "AO", "Administrador", [
        createPermission("micro-a", "AT", "Atualizacao"),
      ]),
    ], {});

    expect(
      findMatchingOtMergeTarget(
        current,
        createPermission("source-macro", "ao", "Administrador"),
        createPermission("source-micro", "at", "Atualizacao"),
      ),
    ).toEqual({ kind: "existing", macroId: "macro-a", microId: "micro-a" });
  });
});

function createTeaDocument(subject: string, activities: TeaDocument["activities"]): TeaDocument {
  return {
    metadata: {
      serviceOrder: "OS",
      phase: "Etapa",
      ticket: "Chamado",
      subject,
      date: "2026-06-22",
      author: "Gabriel",
    },
    overview: "Visao geral",
    activityIntro: "Intro",
    activityImages: [],
    activities,
  };
}

function createTeaActivity(
  id: string,
  title: string,
  blocks = [createTeaTextBlock(`${id}-text`, `Texto de ${title}`)],
  subActivities: TeaDocument["activities"][number]["subActivities"] = [],
): TeaDocument["activities"][number] {
  return {
    id,
    title,
    blocks,
    subActivities,
  };
}

function createTeaSubActivity(
  id: string,
  title: string,
  blocks = [createTeaTextBlock(`${id}-text`, `Texto de ${title}`)],
): TeaDocument["activities"][number]["subActivities"][number] {
  return {
    id,
    title,
    blocks,
  };
}

function createTeaTextBlock(id: string, text: string): TeaDocument["activities"][number]["blocks"][number] {
  return {
    id,
    type: "text",
    text,
  };
}

function createTeaListBlock(
  id: string,
  itemId: string,
  text: string,
): TeaDocument["activities"][number]["blocks"][number] {
  return {
    id,
    type: "list",
    items: [{ id: itemId, text }],
  };
}

function createOtDocument(
  screen: string,
  permissionGroups: OtDocument["permissionGroups"],
  permissionBlocks: OtDocument["permissionBlocks"],
): OtDocument {
  return {
    metadata: {
      screen,
      responsible: "Gabriel",
      date: "2026-06-22",
      environment: "Local",
      author: "Gabriel",
    },
    objective: "Objetivo atual",
    accessSteps: [{ id: "step-a", text: "Passo atual" }],
    permissionGroups,
    permissionBlocks,
  };
}

function createPermissionGroup(
  id: string,
  code: string,
  label: string,
  microPermissions: OtDocument["permissionGroups"][number]["microPermissions"],
): OtDocument["permissionGroups"][number] {
  return {
    id,
    code,
    label,
    selected: true,
    microPermissions,
  };
}

function createPermission(id: string, code: string, label: string) {
  return {
    id,
    code,
    label,
    selected: true,
  };
}

function createTest(
  id: string,
  title: string,
  imageId?: string,
): OtDocument["permissionBlocks"][string]["tests"][number] {
  return {
    id,
    title,
    result: {
      checks: {
        sameBehavior: true,
        possibleIssue: false,
        bothIssue: false,
        newIssue: false,
        errorReport: false,
      },
      observations: "Observacao",
      legacyImages: imageId
        ? [
            {
              id: imageId,
              label: "Legado",
              name: "legado.png",
              dataUrl: "data:image/png;base64,AAA=",
              width: 10,
              height: 10,
            },
          ]
        : [],
      newImages: [],
      errors: [],
    },
  };
}

function idFactory(prefix: string): () => string {
  let index = 0;
  return () => {
    index += 1;
    return `${prefix}-${index}`;
  };
}
