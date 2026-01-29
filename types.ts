
export interface DateRange {
  id: string;
  from: string;
  to: string;
}

export interface Person {
  firstName: string;
  lastName: string;
  title: string;
  titleSuffix: string;
  fullName: string;
  initials: string;
  einrichtung: string;
  kurs: string;
  kursbezeichnung: string;
  strasse: string;
  plzOrt: string;
  email: string;
  telefon: string;
}

export interface TableConfig {
  headerBgColor: string;
  headerTextColor: string;
  borderColor: string;
  borderWidth: number;
}

export interface TemplateField {
  name: string;
  value: string;
}

export enum AppMode {
  ATTENDANCE = 'ATTENDANCE',
  NAMETAGS = 'NAMETAGS',
  ABSENCES = 'ABSENCES',
  SPINE = 'SPINE',
  ROOM = 'ROOM',
  SOL = 'SOL',
  CONTACTS = 'CONTACTS'
}
