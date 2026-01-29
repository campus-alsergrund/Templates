
import { Person } from '../types';

const getAustrianHolidays = (year: number): Set<string> => {
  const holidays = new Set<string>();
  holidays.add(`${year}-01-01`); holidays.add(`${year}-01-06`); holidays.add(`${year}-05-01`);
  holidays.add(`${year}-08-15`); holidays.add(`${year}-10-26`); holidays.add(`${year}-11-01`);
  holidays.add(`${year}-12-08`); holidays.add(`${year}-12-25`); holidays.add(`${year}-12-26`);

  const a = year % 19, b = Math.floor(year / 100), c = year % 100, d = Math.floor(b / 4), e = b % 4,
        f = Math.floor((b + 8) / 25), g = Math.floor((b - f + 1) / 3), h = (19 * a + b - d - g + 15) % 30,
        i = Math.floor(c / 4), k = c % 4, l = (32 + 2 * e + 2 * i - h - k) % 7, m = Math.floor((a + 11 * h + 22 * l) / 451),
        month = Math.floor((h + l - 7 * m + 114) / 31), day = ((h + l - 7 * m + 114) % 31) + 1;
  const easterSunday = new Date(year, month - 1, day);
  const addDays = (base: Date, days: number) => {
    const d = new Date(base); d.setDate(d.getDate() + days);
    return d.toISOString().split('T')[0];
  };
  holidays.add(addDays(easterSunday, 1)); holidays.add(addDays(easterSunday, 39));
  holidays.add(addDays(easterSunday, 50)); holidays.add(addDays(easterSunday, 60));
  return holidays;
};

export const extractKursId = (kursString: string): string => {
  if (!kursString) return "";
  const match = kursString.match(/\d{4,}/); // Sucht nach der ersten Zahl mit mindestens 4 Stellen (z.B. 11061)
  return match ? match[0] : "";
};

export const processNames = (data: any[]): Person[] => {
  return data.map(row => {
    const firstName = (row['Vorname'] || '').trim();
    const lastName = (row['Nachname'] || '').trim();
    const titlePrefix = (row['Titel'] || '').trim();
    const titleSuffix = (row['Titel nachgestellt'] || '').trim();
    const einrichtung = (row['Einrichtung - Kurzbezeichnung'] || row['Einrichtung – Kurzbezeichnung'] || row['Einrichtung'] || '').trim();
    const kurs = (row['Kurs'] || '').trim();
    const kursbezeichnung = (row['Kursbezeichnung'] || '').trim();
    
    const getVal = (row: any, keys: string[]) => {
      for (const k of keys) {
        if (row[k]) return row[k].trim();
        const normalizedK = k.replace(/[–-]/g, '').toLowerCase();
        for (const actualKey in row) {
          if (actualKey.replace(/[–-]/g, '').toLowerCase().includes(normalizedK)) return row[actualKey].trim();
        }
      }
      return '';
    };

    const strasse = getVal(row, ['Privatanschrift – Straße', 'Privatanschrift - Straße', 'Straße', 'Strasse']);
    const plzOrt = getVal(row, ['Privatanschrift – PLZ/Ort', 'Privatanschrift - PLZ/Ort', 'PLZ', 'Ort']);
    const email = getVal(row, ['Privatanschrift – E-Mail Adresse', 'Privatanschrift - E-Mail Adresse', 'E-Mail']);
    const telefon = getVal(row, ['Privatanschrift – Telefon', 'Privatanschrift - Telefon', 'Telefon']);

    let fullName = `${titlePrefix ? titlePrefix + ' ' : ''}${firstName} ${lastName}`.trim();
    if (titleSuffix) fullName += `, ${titleSuffix}`;
    return { firstName, lastName, title: titlePrefix, titleSuffix, fullName, initials: `${firstName.charAt(0)}${lastName.charAt(0)}`, einrichtung, kurs, kursbezeichnung, strasse, plzOrt, email, telefon };
  });
};

export const calculateDatesInRange = (range: {from: string, to: string}): Date[] => {
  if (!range.from || !range.to) return [];
  const start = new Date(range.from), end = new Date(range.to), dates: Date[] = [], year = start.getFullYear();
  const holidays = getAustrianHolidays(year);
  let current = new Date(start);
  while (current <= end) {
    const dow = current.getDay();
    if (dow !== 0 && dow !== 6 && !holidays.has(current.toISOString().split('T')[0])) dates.push(new Date(current));
    current.setDate(current.getDate() + 1);
  }
  return dates;
};

export const calculateDates = (ranges: {from: string, to: string}[]): Date[] => {
  return ranges.flatMap(calculateDatesInRange).sort((a, b) => a.getTime() - b.getTime());
};

export const formatDate = (date: Date): string => date.toLocaleDateString('de-DE', { day: '2-digit', month: '2-digit' });
