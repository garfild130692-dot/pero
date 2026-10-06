// Собственные иконки (простые SVG), без заимствования графики Microsoft.
import React from 'react';

const S = ({ children, size = 18 }) => (
  <svg width={size} height={size} viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="1.8" strokeLinecap="round" strokeLinejoin="round" aria-hidden="true">
    {children}
  </svg>
);

export const Icon = {
  undo: (p) => <S {...p}><path d="M9 14 4 9l5-5" /><path d="M4 9h10a6 6 0 0 1 0 12h-3" /></S>,
  redo: (p) => <S {...p}><path d="m15 14 5-5-5-5" /><path d="M20 9H10a6 6 0 0 0 0 12h3" /></S>,
  save: (p) => <S {...p}><path d="M5 3h11l3 3v15H5z" /><path d="M8 3v6h8V3" /><rect x="8" y="14" width="8" height="7" /></S>,
  open: (p) => <S {...p}><path d="M3 7V5a1 1 0 0 1 1-1h5l2 2h8a1 1 0 0 1 1 1v2" /><path d="M3 7h18l-2 12H5z" /></S>,
  file: (p) => <S {...p}><path d="M6 2h8l5 5v15H6z" /><path d="M14 2v5h5" /></S>,
  saveAs: (p) => <S {...p}><path d="M5 3h11l3 3v7" /><path d="M5 3v18h8" /><path d="M8 3v6h8V3" /><path d="m15 19 5-5" /><path d="m18 14 2 0 0 2" /></S>,
  paste: (p) => <S {...p}><rect x="5" y="4" width="14" height="17" rx="1" /><path d="M9 4V2h6v2" /><path d="M9 11h6M9 15h4" /></S>,
  cut: (p) => <S {...p}><circle cx="6" cy="18" r="3" /><circle cx="18" cy="18" r="3" /><path d="M8 16 18 3M16 16 6 3" /></S>,
  copy: (p) => <S {...p}><rect x="8" y="8" width="12" height="13" rx="1" /><path d="M16 8V4H4v13h4" /></S>,
  clear: (p) => <S {...p}><path d="M5 5h14M12 5l-3 14" /><path d="m15 15 6 6M21 15l-6 6" /></S>,
  bullets: (p) => <S {...p}><circle cx="4" cy="6" r="1.2" fill="currentColor" /><circle cx="4" cy="12" r="1.2" fill="currentColor" /><circle cx="4" cy="18" r="1.2" fill="currentColor" /><path d="M9 6h12M9 12h12M9 18h12" /></S>,
  numbers: (p) => <S {...p}><path d="M3 4h2v5M3 9h3M3 14h3l-3 4h3" strokeWidth="1.5" /><path d="M10 6h11M10 12h11M10 18h11" /></S>,
  outdent: (p) => <S {...p}><path d="M3 4h18M11 9h10M11 15h10M3 20h18" /><path d="m7 9-4 3 4 3" /></S>,
  indent: (p) => <S {...p}><path d="M3 4h18M11 9h10M11 15h10M3 20h18" /><path d="m3 9 4 3-4 3" /></S>,
  search: (p) => <S {...p}><circle cx="10.5" cy="10.5" r="6" /><path d="m15 15 6 6" /></S>,
  replace: (p) => <S {...p}><path d="M4 7h11l-3-3M20 17H9l3 3" /><path d="M4 7v3M20 17v-3" /></S>,
  comment: (p) => <S {...p}><path d="M4 5h16v11H9l-5 4z" /><path d="M12 8v5M9.5 10.5h5" /></S>,
  comments: (p) => <S {...p}><path d="M3 4h13v9H7l-4 3z" /><path d="M8 15v2h9l4 3V9h-3" /></S>,
  print: (p) => <S {...p}><path d="M7 9V3h10v6" /><rect x="3" y="9" width="18" height="8" rx="1.5" /><path d="M7 14h10v7H7z" /></S>,
  pdf: (p) => <S {...p}><path d="M6 2h8l5 5v15H6z" /><path d="M14 2v5h5" /><path d="M8.5 17v-4h1.3a1.2 1.2 0 0 1 0 2.4H8.5M12.5 13v4h1a2 2 0 0 0 0-4zM17 13h-1.8v4M15.2 15h1.4" strokeWidth="1.1" /></S>,
  picture: (p) => <S {...p}><rect x="3" y="4" width="18" height="16" rx="1.5" /><circle cx="9" cy="10" r="2" /><path d="m21 17-5-5-9 8" /></S>,
  lineSpacing: (p) => <S {...p}><path d="M11 5h10M11 10h10M11 15h10M11 20h10" /><path d="m3 7 3-3 3 3M6 4v16M3 17l3 3 3-3" /></S>,
  launcher: (p) => <S {...p}><path d="M5 5h9M5 5v9" /><path d="m9 9 10 10M19 12v7h-7" /></S>,
  alignLeft: (p) => <S {...p}><path d="M3 5h18M3 10h12M3 15h18M3 20h12" /></S>,
  alignCenter: (p) => <S {...p}><path d="M3 5h18M6 10h12M3 15h18M6 20h12" /></S>,
  alignRight: (p) => <S {...p}><path d="M3 5h18M9 10h12M3 15h18M9 20h12" /></S>,
  alignJustify: (p) => <S {...p}><path d="M3 5h18M3 10h18M3 15h18M3 20h18" /></S>,
  table: (p) => <S {...p}><rect x="3" y="4" width="18" height="16" rx="1" /><path d="M3 9h18M3 14h18M9 4v16M15 4v16" /></S>,
  link: (p) => <S {...p}><path d="M10 14a4 4 0 0 0 5.7 0l3-3a4 4 0 0 0-5.7-5.7l-1 1" /><path d="M14 10a4 4 0 0 0-5.7 0l-3 3a4 4 0 0 0 5.7 5.7l1-1" /></S>,
  unlink: (p) => <S {...p}><path d="M15 12l3.7-3.7a4 4 0 0 0-5.7-5.7L11 4.6" /><path d="M9 12l-3.7 3.7a4 4 0 0 0 5.7 5.7l1.9-1.9" /><path d="M4 4l16 16" /></S>,
  lineBreak: (p) => <S {...p}><path d="M20 5v6a3 3 0 0 1-3 3H5" /><path d="m9 10-4 4 4 4" /></S>,
  rowAbove: (p) => <S {...p}><rect x="3" y="11" width="18" height="10" rx="1" /><path d="M3 16h18M12 2v6M9 5h6" /></S>,
  rowBelow: (p) => <S {...p}><rect x="3" y="3" width="18" height="10" rx="1" /><path d="M3 8h18M12 16v6M9 19h6" /></S>,
  colLeft: (p) => <S {...p}><rect x="11" y="3" width="10" height="18" rx="1" /><path d="M16 3v18M2 12h6M5 9v6" /></S>,
  colRight: (p) => <S {...p}><rect x="3" y="3" width="10" height="18" rx="1" /><path d="M8 3v18M16 12h6M19 9v6" /></S>,
  delRow: (p) => <S {...p}><rect x="3" y="8" width="18" height="8" rx="1" /><path d="m9 2 6 4M15 2 9 6" /></S>,
  delCol: (p) => <S {...p}><rect x="8" y="3" width="8" height="18" rx="1" /><path d="m18 9 4 6M22 9l-4 6" /></S>,
  delTable: (p) => <S {...p}><rect x="3" y="4" width="13" height="12" rx="1" /><path d="M3 10h13M9 4v12" /><path d="m16 16 5 5M21 16l-5 5" /></S>,
  merge: (p) => <S {...p}><rect x="3" y="5" width="18" height="14" rx="1" /><path d="m7 12 3-2v4zM17 12l-3-2v4z" fill="currentColor" /></S>,
  split: (p) => <S {...p}><rect x="3" y="5" width="18" height="14" rx="1" /><path d="M12 5v14" /></S>,
};
