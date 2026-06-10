import { Loader, Text } from "@mantine/core";

export type LoadingFeedbackProps = {
  label: string;
  detail?: string;
  variant?: "global" | "inline" | "compact";
  className?: string;
  id?: string;
};

export function LoadingFeedback({
  label,
  detail,
  variant = "inline",
  className,
  id,
}: LoadingFeedbackProps) {
  const classNames = [
    "loadingFeedback",
    `loadingFeedback--${variant}`,
    className,
  ]
    .filter(Boolean)
    .join(" ");

  return (
    <div
      id={id}
      className={classNames}
      role="status"
      aria-live="polite"
      aria-atomic="true"
    >
      <Loader size={variant === "compact" ? "xs" : "sm"} type="oval" />
      <div className="loadingFeedbackCopy">
        <Text fw={800} size={variant === "global" ? "sm" : "xs"} className="loadingFeedbackLabel">
          {label}
        </Text>
        {detail ? (
          <Text size="xs" c="dimmed" className="loadingFeedbackDetail">
            {detail}
          </Text>
        ) : null}
      </div>
    </div>
  );
}
