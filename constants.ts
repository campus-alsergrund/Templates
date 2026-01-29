
export const DEFAULT_TABLE_CONFIG = {
  headerBgColor: '#D8E8C5',
  headerTextColor: '#000000',
  borderColor: '#000000',
  borderWidth: 2, // pt
};

export const TITLES_TO_EXCLUDE = [
  'dr.', 'mag.', 'bsc.', 'msc.', 'dipl.-ing.', 'prof.', 'dr. med.', 'phd', 'mba', 'b.a.', 'm.a.'
];

export const REQUIRED_TEMPLATES = [
  { id: 'attendance', label: 'Anwesenheitsliste', filename: 'Anwesenheitsliste.docx' },
  { id: 'nametags', label: 'Namensschilder', filename: 'Namensschilder.docx' },
  { id: 'absences', label: 'Absenzenliste', filename: 'Absenzenliste.docx' },
  { id: 'spine', label: 'Ordnerrücken', filename: 'Ordnerruecken.docx' },
  { id: 'room', label: 'Raumbeschriftung', filename: 'Raumbeschriftung.docx' },
  { id: 'sol', label: 'SOL-Liste', filename: 'SOL_Liste.docx' },
  { id: 'contacts', label: 'Kontaktdaten', filename: 'Kontaktdaten.docx' }
];
