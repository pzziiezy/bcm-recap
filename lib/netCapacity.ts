/**
 * Net Capacity = P% × O (Piece), rounded to 2 decimals.
 * Shared by download.worker.ts (RECAP F-J/N-Q patch) and lib/minorReport.ts —
 * kept dependency-free (no xlsx import) so it stays cheap to pull into the Worker bundle.
 */
export function computeNetCapacity(pctStr: string | undefined, pieceStr: string | undefined): number | null {
  const pctNum = parseFloat(pctStr ?? "0") || 0;
  const pieceNum = parseFloat(pieceStr ?? "0") || 0;
  if (pctNum > 0 && pieceNum > 0) {
    return Math.round((pctNum / 100) * pieceNum * 100) / 100;
  }
  return null;
}
