import { memo } from "react";
import type { CSSProperties } from "react";
import type {
  DocxPreviewBlock,
  DocxPreviewCell,
  DocxPreviewModel,
  DocxPreviewParagraph,
  DocxPreviewRun,
  DocxPreviewTable,
} from "./docxPreviewModel";

export const DocxPreview = memo(function DocxPreview({ model }: { model: DocxPreviewModel }) {
  return (
    <section
      id={model.sectionId}
      className="docxPreviewSection"
      aria-label={model.title}
      tabIndex={-1}
    >
      <div className="docxPreviewShell">
        <div
          className={`docxPreviewPage docxPreviewPage--${model.kind}`}
          role="document"
          aria-label={model.title}
        >
          {/* Blocos principais carregam id do modelo; o indice entra so como desempate para ids gerados por conteudo repetido. */}
          {model.blocks.map((block, index) => renderBlock(block, index))}
        </div>
      </div>
    </section>
  );
});

function renderBlock(block: DocxPreviewBlock, blockIndex: number) {
  if (block.type === "paragraph") {
    return renderParagraph(block, blockIndex);
  }

  if (block.type === "table") {
    return renderTable(block, blockIndex);
  }

  if (block.type === "list") {
    return (
      <ul key={getBlockKey(block, blockIndex)} className="docxPreviewList">
        {/* Itens de lista nao tem id proprio; tipo, posicao e prefixo do texto estabilizam insercoes/remocoes proximas. */}
        {block.items.map((item, itemIndex) => (
          <li key={`list-item-${block.id}-${itemIndex}-${getRunsTextPrefix(item, 12)}`}>
            {renderRuns(item, `list-${blockIndex}-item-${itemIndex}`)}
          </li>
        ))}
      </ul>
    );
  }

  return (
    <div
      key={getBlockKey(block, blockIndex)}
      className={`docxPreviewImage docxPreviewImage--${block.variant}`}
    >
      <img
        src={block.src}
        alt={block.alt}
        width={block.width}
        height={block.height}
        style={{ width: block.width, height: block.height }}
      />
    </div>
  );
}

function renderParagraph(block: DocxPreviewParagraph, paragraphIndex: number) {
  const style = block.fill ? ({ backgroundColor: `#${block.fill}` } as CSSProperties) : undefined;

  return (
    <p
      id={block.anchorId}
      key={`p-${paragraphIndex}-${block.variant}-${getRunsTextPrefix(block.runs, 12)}-${block.id}`}
      className={`docxPreviewParagraph docxPreviewParagraph--${block.variant} ${
        block.alignment === "center" ? "docxPreviewParagraph--center" : ""
      }`}
      style={style}
    >
      {renderRuns(block.runs, `p-${paragraphIndex}`)}
    </p>
  );
}

function renderTable(block: DocxPreviewTable, tableIndex: number) {
  return (
    <table
      id={block.anchorId}
      key={getBlockKey(block, tableIndex)}
      className={`docxPreviewTable docxPreviewTable--${block.variant}`}
    >
      <colgroup>
        {/* Colunas nao sao reordenaveis no preview; largura + posicao e suficiente como chave composta. */}
        {block.columnWidths.map((width, widthIndex) => (
          <col key={`${block.id}-col-${widthIndex}-${width}`} style={{ width: `${width}%` }} />
        ))}
      </colgroup>
      <tbody>
        {/* Linhas e celulas nao carregam id; usamos tabela + tipo + posicao para preservar reconciliacao local. */}
        {block.rows.map((row, rowIndex) => (
          <tr key={`${block.id}-row-${rowIndex}-${row.length}`}>
            {row.map((cell, cellIndex) =>
              renderCell(cell, block.variant, tableIndex, rowIndex, cellIndex),
            )}
          </tr>
        ))}
      </tbody>
    </table>
  );
}

function renderCell(
  cell: DocxPreviewCell,
  variant: DocxPreviewTable["variant"],
  tableIndex: number,
  rowIndex: number,
  cellIndex: number,
) {
  const style = cell.fill ? ({ backgroundColor: `#${cell.fill}` } as CSSProperties) : undefined;
  const cellKey = `cell-${tableIndex}-${rowIndex}-${cellIndex}-${cell.columnSpan ?? 1}-${cell.fill ?? "none"}`;

  return (
    <td
      key={cellKey}
      colSpan={cell.columnSpan}
      className={`docxPreviewCell docxPreviewCell--${variant} ${
        cell.bold ? "docxPreviewCell--label" : ""
      } ${cell.alignment === "center" ? "docxPreviewCell--center" : ""}`}
      style={style}
    >
      {/* Paragrafos de celula vem como runs sem id; posicao da celula + prefixo do texto evita key=index puro. */}
      {cell.paragraphs.map((paragraph, paragraphIndex) => (
        <p
          key={`${cellKey}-paragraph-${paragraphIndex}-${getRunsTextPrefix(paragraph, 12)}`}
          className="docxPreviewCellParagraph"
        >
          {renderRuns(paragraph, `${cellKey}-p-${paragraphIndex}`)}
        </p>
      ))}
    </td>
  );
}

function renderRuns(runs: DocxPreviewRun[], keyContext: string) {
  // Runs tambem nao tem id no modelo; contexto do paragrafo + posicao + prefixo do texto e estavel o bastante para o preview.
  return runs.map((run, runIndex) => (
    <span
      key={`run-${keyContext}-${runIndex}-${run.text.slice(0, 8)}-${run.bold ? "b" : "n"}-${run.italic ? "i" : "n"}-${run.underline ? "u" : "n"}`}
      className={[
        "docxPreviewRun",
        run.bold ? "docxPreviewRun--bold" : "",
        run.italic ? "docxPreviewRun--italic" : "",
        run.underline ? "docxPreviewRun--underline" : "",
      ]
        .filter(Boolean)
        .join(" ")}
    >
      {run.text}
    </span>
  ));
}

function getBlockKey(block: DocxPreviewBlock, blockIndex: number): string {
  if (block.type === "paragraph") {
    return `p-${blockIndex}-${block.variant}-${getRunsTextPrefix(block.runs, 12)}-${block.id}`;
  }

  if (block.type === "table") {
    return `table-${blockIndex}-${block.variant}-${block.rows.length}-${block.id}`;
  }

  if (block.type === "list") {
    return `list-${blockIndex}-${getRunsTextPrefix(block.items[0] ?? [], 12)}-${block.id}`;
  }

  return `image-${blockIndex}-${block.variant}-${block.alt.slice(0, 12)}-${block.id}`;
}

function getRunsTextPrefix(runs: DocxPreviewRun[], length: number): string {
  return runs
    .map((run) => run.text)
    .join("")
    .trim()
    .slice(0, length);
}
