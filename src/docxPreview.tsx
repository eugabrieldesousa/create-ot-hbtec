import type { CSSProperties } from "react";
import type {
  DocxPreviewBlock,
  DocxPreviewCell,
  DocxPreviewModel,
  DocxPreviewParagraph,
  DocxPreviewRun,
  DocxPreviewTable,
} from "./docxPreviewModel";

export function DocxPreview({ model }: { model: DocxPreviewModel }) {
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
          {model.blocks.map((block, index) => renderBlock(block, index))}
        </div>
      </div>
    </section>
  );
}

function renderBlock(block: DocxPreviewBlock, index: number) {
  if (block.type === "paragraph") {
    return renderParagraph(block, index);
  }

  if (block.type === "table") {
    return renderTable(block, index);
  }

  if (block.type === "list") {
    return (
      <ul key={`${block.id}-${index}`} className="docxPreviewList">
        {block.items.map((item, itemIndex) => (
          <li key={`${block.id}-item-${itemIndex}`}>{renderRuns(item)}</li>
        ))}
      </ul>
    );
  }

  return (
    <div
      key={`${block.id}-${index}`}
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

function renderParagraph(block: DocxPreviewParagraph, index: number) {
  const style = block.fill ? ({ backgroundColor: `#${block.fill}` } as CSSProperties) : undefined;

  return (
    <p
      id={block.anchorId}
      key={`${block.id}-${index}`}
      className={`docxPreviewParagraph docxPreviewParagraph--${block.variant} ${
        block.alignment === "center" ? "docxPreviewParagraph--center" : ""
      }`}
      style={style}
    >
      {renderRuns(block.runs)}
    </p>
  );
}

function renderTable(block: DocxPreviewTable, index: number) {
  return (
    <table
      key={`${block.id}-${index}`}
      className={`docxPreviewTable docxPreviewTable--${block.variant}`}
    >
      <colgroup>
        {block.columnWidths.map((width, widthIndex) => (
          <col key={`${block.id}-col-${widthIndex}`} style={{ width: `${width}%` }} />
        ))}
      </colgroup>
      <tbody>
        {block.rows.map((row, rowIndex) => (
          <tr key={`${block.id}-row-${rowIndex}`}>
            {row.map((cell, cellIndex) => renderCell(cell, block.variant, cellIndex))}
          </tr>
        ))}
      </tbody>
    </table>
  );
}

function renderCell(cell: DocxPreviewCell, variant: DocxPreviewTable["variant"], index: number) {
  const style = cell.fill ? ({ backgroundColor: `#${cell.fill}` } as CSSProperties) : undefined;

  return (
    <td
      key={`cell-${index}`}
      colSpan={cell.columnSpan}
      className={`docxPreviewCell docxPreviewCell--${variant} ${
        cell.bold ? "docxPreviewCell--label" : ""
      } ${cell.alignment === "center" ? "docxPreviewCell--center" : ""}`}
      style={style}
    >
      {cell.paragraphs.map((paragraph, paragraphIndex) => (
        <p key={`cell-paragraph-${paragraphIndex}`} className="docxPreviewCellParagraph">
          {renderRuns(paragraph)}
        </p>
      ))}
    </td>
  );
}

function renderRuns(runs: DocxPreviewRun[]) {
  return runs.map((run, index) => (
    <span
      key={`${run.text}-${index}`}
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
