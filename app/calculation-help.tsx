import type { ReactNode } from "react";

export function CalculationHelp({ source, children, calculated = true }: { source?: string; children: ReactNode; calculated?: boolean }) {
  const exported = Boolean((globalThis as any).__WTDC_SNAPSHOT__);
  const label = calculated ? "How this is calculated" : "About this information";
  return <details className="calculation-help"><summary aria-label={label} title={label}><span className="info-symbol" aria-hidden="true">i</span><span className="calculation-label">{label}</span></summary><div className="calculation-body">{source && !exported && <p><strong>Source:</strong> {source}</p>}<p>{children}</p></div></details>;
}
