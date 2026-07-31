// Theme tokens for the grid renderer. Light default; dark can override.

export interface GridTheme {
  background: string;
  gridLine: string;
  headerBackground: string;
  headerText: string;
  headerHighlight: string;
  headerHighlightText: string;
  cellText: string;
  cellErrorText: string;
  selectionFill: string;
  selectionBorder: string;
  activeCellBorder: string;
  frozenDivider: string;
  scrollbarTrack: string;
  scrollbarThumb: string;
  font: string;
  headerFont: string;
}

export const lightTheme: GridTheme = {
  background: "#ffffff",
  gridLine: "#e2e4e8",
  headerBackground: "#f7f8fa",
  headerText: "#5f6368",
  headerHighlight: "#d3e3fd",
  headerHighlightText: "#0b57d0",
  cellText: "#1f2329",
  cellErrorText: "#b3261e",
  selectionFill: "rgba(11, 87, 208, 0.10)",
  selectionBorder: "rgba(11, 87, 208, 0.45)",
  activeCellBorder: "#0b57d0",
  frozenDivider: "#9aa0a6",
  scrollbarTrack: "rgba(0, 0, 0, 0.04)",
  scrollbarThumb: "rgba(0, 0, 0, 0.25)",
  font: '13px -apple-system, "Segoe UI", Roboto, "Helvetica Neue", Arial, sans-serif',
  headerFont: '11px -apple-system, "Segoe UI", Roboto, "Helvetica Neue", Arial, sans-serif',
};

export const darkTheme: GridTheme = {
  background: "#1f2329",
  gridLine: "#33373d",
  headerBackground: "#26292f",
  headerText: "#9aa0a6",
  headerHighlight: "#004a77",
  headerHighlightText: "#c2e7ff",
  cellText: "#e3e5e8",
  cellErrorText: "#f2b8b5",
  selectionFill: "rgba(168, 199, 250, 0.16)",
  selectionBorder: "rgba(168, 199, 250, 0.5)",
  activeCellBorder: "#a8c7fa",
  frozenDivider: "#5f6368",
  scrollbarTrack: "rgba(255, 255, 255, 0.05)",
  scrollbarThumb: "rgba(255, 255, 255, 0.28)",
  font: '13px -apple-system, "Segoe UI", Roboto, "Helvetica Neue", Arial, sans-serif',
  headerFont: '11px -apple-system, "Segoe UI", Roboto, "Helvetica Neue", Arial, sans-serif',
};
