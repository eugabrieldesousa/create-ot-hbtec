import type { TeaDocument } from "./types";

export function createDefaultTeaDocument(): TeaDocument {
  const today = new Date().toISOString().slice(0, 10);

  return {
    metadata: {
      serviceOrder: "",
      phase: "Etapa 5",
      ticket: "",
      subject: "Telas - Novo Layout",
      date: today,
      author: "Gabriel Sousa",
    },
    overview: "",
    activityIntro: "A seguir serao apresentadas, a nova interface e as suas funcionalidades:",
    activityImages: [],
    activities: [
      {
        id: "tea-activity-default",
        title: "",
        blocks: [],
        subActivities: [],
      },
    ],
  };
}
