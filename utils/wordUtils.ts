
import { Person, TableConfig, TemplateField, AppMode } from '../types';
import { formatDate, calculateDatesInRange } from './csvUtils';

const cleanAndEscape = (unsafe: any): string => {
  if (unsafe === undefined || unsafe === null) return "";
  const s = String(unsafe);
  const clean = s.replace(/[\x00-\x08\x0B\x0C\x0E-\x1F]/g, "");
  return clean.replace(/[<>&'"]/g, (c) => {
    switch (c) {
      case '<': return '&lt;'; case '>': return '&gt;'; case '&': return '&amp;'; case '\'': return '&apos;'; case '"': return '&quot;';
      default: return c;
    }
  });
};

const LUCIDA = '<w:rPr><w:rFonts w:ascii="Lucida Sans Unicode" w:hAnsi="Lucida Sans Unicode" w:cs="Lucida Sans Unicode"/></w:rPr>';
const VCENTER = '<w:tcPr><w:vAlign w:val="center"/></w:tcPr>';

export const generateFinalDocx = async (
  templateBuffer: ArrayBuffer,
  persons: Person[],
  ranges: any[],
  fields: TemplateField[],
  config: TableConfig,
  showInitials: boolean,
  mode: AppMode
): Promise<Blob> => {
  const PizZip = (window as any).PizZip;
  const Docxtemplater = (window as any).docxtemplater;

  if (mode === AppMode.NAMETAGS) {
    let combinedBodyXml = "";
    for (let i = 0; i < persons.length; i++) {
      const p = persons[i];
      const zip = new PizZip(templateBuffer);
      const doc = new Docxtemplater(zip, { paragraphLoop: true, linebreaks: true, delimiters: { start: '{{', end: '}}' } });
      
      const personData: Record<string, any> = {};
      fields.forEach(f => {
        const n = f.name.toLowerCase();
        if (n === 'vorname') personData[f.name] = cleanAndEscape(p.firstName);
        else if (n === 'nachname') personData[f.name] = cleanAndEscape(p.lastName);
        else if (n === 'titel') personData[f.name] = cleanAndEscape(p.title);
        else if (n.includes('komma')) personData[f.name] = p.titleSuffix ? `, ${cleanAndEscape(p.titleSuffix)}` : "";
        else if (n.includes('nachgestellt')) personData[f.name] = cleanAndEscape(p.titleSuffix);
        else if (n === 'einrichtung') personData[f.name] = cleanAndEscape(p.einrichtung);
        else personData[f.name] = cleanAndEscape(f.value);
      });
      
      doc.render(personData);
      const xml = doc.getZip().file("word/document.xml").asText();
      const bodyMatch = xml.match(/<w:body>(.*?)<\/w:body>/s);
      if (bodyMatch) {
        let bodyContent = bodyMatch[1];
        bodyContent = bodyContent.replace(/<w:sectPr>.*?<\/w:sectPr>/gs, "");
        combinedBodyXml += bodyContent;
        if (i < persons.length - 1) {
          combinedBodyXml += '<w:p><w:r><w:br w:type="page"/></w:r></w:p>';
        }
      }
    }
    const finalZip = new PizZip(templateBuffer);
    let finalXml = finalZip.file("word/document.xml").asText();
    finalXml = finalXml.replace(/<w:body>.*?<\/w:body>/s, `<w:body>${combinedBodyXml}</w:body>`);
    finalZip.file("word/document.xml", finalXml);
    return finalZip.generate({ type: "blob", compression: "DEFLATE" });
  }

  const zip = new PizZip(templateBuffer);
  const data: Record<string, any> = {
    maintable: '{{maintable}}'
  };
  fields.forEach(f => { data[f.name] = cleanAndEscape(f.value); });
  
  if (persons.length > 0) {
    data['Kurs'] = data['Kurs'] || cleanAndEscape(persons[0].kurs);
    data['Kursbezeichnung'] = data['Kursbezeichnung'] || cleanAndEscape(persons[0].kursbezeichnung);
    data['Einrichtung'] = data['Einrichtung'] || cleanAndEscape(persons[0].einrichtung);
  }

  const doc = new Docxtemplater(zip, { 
    paragraphLoop: true, 
    linebreaks: true, 
    delimiters: { start: '{{', end: '}}' },
    nullGetter() { return ""; } 
  });

  doc.render(data);
  const updatedZip = doc.getZip();
  let documentXml = updatedZip.file("word/document.xml").asText();
  
  let tableXml = '';
  switch(mode) {
    case AppMode.ATTENDANCE: tableXml = generateAttendanceByRange(persons, ranges, config); break;
    case AppMode.SOL: tableXml = generateSolTableXmlContent(persons, config); break;
    case AppMode.CONTACTS: tableXml = generateContactsTableXmlContent(persons, config); break;
    case AppMode.ABSENCES: tableXml = generateAbsencesTableXmlContent(persons, config); break;
    default: tableXml = ''; 
  }

  const marker = '{{maintable}}';
  const markerIdx = documentXml.indexOf(marker);
  if (markerIdx !== -1 && tableXml) {
    const actualPStart = Math.max(documentXml.lastIndexOf('<w:p ', markerIdx), documentXml.lastIndexOf('<w:p>', markerIdx));
    const pEndIdx = documentXml.indexOf('</w:p>', markerIdx);
    if (actualPStart !== -1 && pEndIdx !== -1) {
      documentXml = documentXml.substring(0, actualPStart) + tableXml + documentXml.substring(pEndIdx + 6);
    }
  }

  updatedZip.file("word/document.xml", documentXml);
  return updatedZip.generate({ type: "blob", compression: "DEFLATE" });
};

const getTblPr = (config: TableConfig) => {
  const c = config.borderColor.replace('#', ''), s = Math.max(2, config.borderWidth * 2);
  return `<w:tblPr><w:tblStyle w:val="TableGrid"/><w:tblW w:w="5000" w:type="pct"/><w:jc w:val="center"/><w:tblBorders><w:top w:val="single" w:sz="${s}" w:color="${c}"/><w:left w:val="single" w:sz="${s}" w:color="${c}"/><w:bottom w:val="single" w:sz="${s}" w:color="${c}"/><w:right w:val="single" w:sz="${s}" w:color="${c}"/><w:insideH w:val="single" w:sz="${s}" w:color="${c}"/><w:insideV w:val="single" w:sz="${s}" w:color="${c}"/></w:tblBorders></w:tblPr>`;
};

const generateAttendanceByRange = (persons: Person[], ranges: any[], config: TableConfig): string => {
  let xml = '';
  const hBg = config.headerBgColor.replace('#', ''), hText = config.headerTextColor.replace('#', '');
  ranges.forEach((range, rIdx) => {
    const dates = calculateDatesInRange(range);
    if (dates.length === 0) return;
    const CHUNK = 12;
    for (let i = 0; i < dates.length; i += CHUNK) {
      const chunk = dates.slice(i, i + CHUNK);
      xml += `<w:tbl>${getTblPr(config)}<w:tblGrid><w:gridCol w:w="500"/><w:gridCol w:w="3600"/>${chunk.map(() => '<w:gridCol w:w="450"/>').join('')}</w:tblGrid>
        <w:tr><w:trPr><w:trHeight w:val="600"/></w:trPr>
          <w:tc>${VCENTER}<w:tcPr><w:shd w:val="clear" w:fill="${hBg}"/></w:tcPr><w:p><w:pPr><w:jc w:val="center"/></w:pPr><w:r>${LUCIDA}<w:rPr><w:color w:val="${hText}"/><w:b/></w:rPr><w:t>HZ</w:t></w:r></w:p></w:tc>
          <w:tc>${VCENTER}<w:tcPr><w:shd w:val="clear" w:fill="${hBg}"/></w:tcPr><w:p><w:pPr><w:jc w:val="center"/></w:pPr><w:r>${LUCIDA}<w:rPr><w:color w:val="${hText}"/><w:b/></w:rPr><w:t>Teilnehmer/in</w:t></w:r></w:p></w:tc>
          ${chunk.map(d => `<w:tc>${VCENTER}<w:tcPr><w:shd w:val="clear" w:fill="${hBg}"/></w:tcPr><w:p><w:pPr><w:jc w:val="center"/></w:pPr><w:r>${LUCIDA}<w:rPr><w:color w:val="${hText}"/><w:b/><w:sz w:val="14"/></w:rPr><w:t>${formatDate(d)}</w:t></w:r></w:p></w:tc>`).join('')}
        </w:tr>`;
      persons.forEach(p => {
        xml += `<w:tr><w:trPr><w:trHeight w:val="450"/></w:trPr>
          <w:tc>${VCENTER}<w:p><w:pPr><w:jc w:val="center"/></w:pPr><w:r>${LUCIDA}<w:t>${cleanAndEscape(p.initials)}</w:t></w:r></w:p></w:tc>
          <w:tc>${VCENTER}<w:p><w:pPr><w:jc w:val="left"/><w:spacing w:before="60" w:after="60"/></w:pPr><w:r>${LUCIDA}<w:t>${cleanAndEscape(p.fullName)}</w:t></w:r></w:p></w:tc>
          ${chunk.map(() => `<w:tc>${VCENTER}<w:p/></w:tc>`).join('')}
        </w:tr>`;
      });
      xml += `</w:tbl>`;
      if (i + CHUNK < dates.length || rIdx < ranges.length - 1) xml += `<w:p><w:r><w:br w:type="page"/></w:r></w:p>`;
    }
  });
  return xml;
};

const generateSolTableXmlContent = (persons: Person[], config: TableConfig): string => {
  const hBg = config.headerBgColor.replace('#', ''), hText = config.headerTextColor.replace('#', '');
  let xml = `<w:tbl>${getTblPr(config)}<w:tblGrid><w:gridCol w:w="4200"/><w:gridCol w:w="2800"/><w:gridCol w:w="3000"/></w:tblGrid>
    <w:tr><w:trPr><w:trHeight w:val="800"/></w:trPr>
      <w:tc>${VCENTER}<w:tcPr><w:shd w:val="clear" w:fill="${hBg}"/></w:tcPr><w:p><w:pPr><w:jc w:val="center"/></w:pPr><w:r>${LUCIDA}<w:rPr><w:color w:val="${hText}"/><w:b/></w:rPr><w:t>Name</w:t></w:r></w:p></w:tc>
      <w:tc>${VCENTER}<w:tcPr><w:shd w:val="clear" w:fill="${hBg}"/></w:tcPr><w:p><w:pPr><w:jc w:val="center"/></w:pPr><w:r>${LUCIDA}<w:rPr><w:color w:val="${hText}"/><w:b/></w:rPr><w:t>Aufenthaltsort</w:t></w:r></w:p></w:tc>
      <w:tc>${VCENTER}<w:tcPr><w:shd w:val="clear" w:fill="${hBg}"/></w:tcPr><w:p><w:pPr><w:jc w:val="center"/></w:pPr><w:r>${LUCIDA}<w:rPr><w:color w:val="${hText}"/><w:b/></w:rPr><w:t>Unterschrift</w:t></w:r></w:p></w:tc>
    </w:tr>`;
  persons.forEach(p => {
    xml += `<w:tr><w:trPr><w:trHeight w:val="500"/></w:trPr>
      <w:tc>${VCENTER}<w:p><w:pPr><w:spacing w:before="120" w:after="120"/></w:pPr><w:r>${LUCIDA}<w:t>${cleanAndEscape(p.fullName)}</w:t></w:r></w:p></w:tc>
      <w:tc>${VCENTER}<w:p/></w:tc><w:tc>${VCENTER}<w:p/></w:tc>
    </w:tr>`;
  });
  return xml + `</w:tbl>`;
};

const generateContactsTableXmlContent = (persons: Person[], config: TableConfig): string => {
  const hBg = config.headerBgColor.replace('#', ''), hText = config.headerTextColor.replace('#', '');
  // Layout angepasst auf 5 Spalten, Name kombiniert wie in SOL-Liste
  let xml = `<w:tbl>${getTblPr(config)}<w:tblGrid><w:gridCol w:w="3200"/><w:gridCol w:w="1800"/><w:gridCol w:w="1500"/><w:gridCol w:w="2000"/><w:gridCol w:w="1500"/></w:tblGrid>
    <w:tr><w:trPr><w:trHeight w:val="600"/></w:trPr>
      ${['Name','Straße','PLZ / Ort','E-Mail','Telefon'].map(h => `<w:tc>${VCENTER}<w:tcPr><w:shd w:val="clear" w:fill="${hBg}"/></w:tcPr><w:p><w:pPr><w:jc w:val="center"/></w:pPr><w:r>${LUCIDA}<w:rPr><w:color w:val="${hText}"/><w:b/></w:rPr><w:t>${h}</w:t></w:r></w:p></w:tc>`).join('')}
    </w:tr>`;
  persons.forEach(p => {
    xml += `<w:tr><w:trPr><w:trHeight w:val="450"/></w:trPr>
      <w:tc>${VCENTER}<w:p><w:r>${LUCIDA}<w:t>${cleanAndEscape(p.fullName)}</w:t></w:r></w:p></w:tc>
      <w:tc>${VCENTER}<w:p><w:r>${LUCIDA}<w:t>${cleanAndEscape(p.strasse)}</w:t></w:r></w:p></w:tc>
      <w:tc>${VCENTER}<w:p><w:r>${LUCIDA}<w:t>${cleanAndEscape(p.plzOrt)}</w:t></w:r></w:p></w:tc>
      <w:tc>${VCENTER}<w:p><w:r>${LUCIDA}<w:t>${cleanAndEscape(p.email)}</w:t></w:r></w:p></w:tc>
      <w:tc>${VCENTER}<w:p><w:r>${LUCIDA}<w:t>${cleanAndEscape(p.telefon)}</w:t></w:r></w:p></w:tc>
    </w:tr>`;
  });
  return xml + `</w:tbl>`;
};

const generateAbsencesTableXmlContent = (persons: Person[], config: TableConfig): string => {
  const hBg = config.headerBgColor.replace('#', ''), hText = config.headerTextColor.replace('#', '');
  let xml = `<w:tbl>${getTblPr(config)}<w:tblGrid><w:gridCol w:w="3000"/><w:gridCol w:w="800"/><w:gridCol w:w="3200"/><w:gridCol w:w="3000"/></w:tblGrid>
    <w:tr><w:trPr><w:trHeight w:val="600"/></w:trPr>
      <w:tc>${VCENTER}<w:tcPr><w:shd w:val="clear" w:fill="${hBg}"/></w:tcPr><w:p><w:pPr><w:jc w:val="center"/></w:pPr><w:r>${LUCIDA}<w:rPr><w:color w:val="${hText}"/><w:b/></w:rPr><w:t>Teilnehmer/in</w:t></w:r></w:p></w:tc>
      <w:tc>${VCENTER}<w:tcPr><w:shd w:val="clear" w:fill="${hBg}"/></w:tcPr><w:p><w:pPr><w:jc w:val="center"/></w:pPr><w:r>${LUCIDA}<w:rPr><w:color w:val="${hText}"/><w:b/></w:rPr><w:t>Abk.</w:t></w:r></w:p></w:tc>
      <w:tc>${VCENTER}<w:tcPr><w:shd w:val="clear" w:fill="${hBg}"/></w:tcPr><w:p><w:pPr><w:jc w:val="center"/></w:pPr><w:r>${LUCIDA}<w:rPr><w:color w:val="${hText}"/><w:b/></w:rPr><w:t>Grund der Abwesenheit</w:t></w:r></w:p></w:tc>
      <w:tc>${VCENTER}<w:tcPr><w:shd w:val="clear" w:fill="${hBg}"/></w:tcPr><w:p><w:pPr><w:jc w:val="center"/></w:pPr><w:r>${LUCIDA}<w:rPr><w:color w:val="${hText}"/><w:b/></w:rPr><w:t>Unterschrift Trainer/in</w:t></w:r></w:p></w:tc>
    </w:tr>`;
  persons.forEach(p => {
    xml += `<w:tr><w:trPr><w:trHeight w:val="450"/></w:trPr>
      <w:tc>${VCENTER}<w:p><w:r>${LUCIDA}<w:t>${cleanAndEscape(p.fullName)}</w:t></w:r></w:p></w:tc>
      <w:tc>${VCENTER}<w:p><w:pPr><w:jc w:val="center"/></w:pPr><w:r>${LUCIDA}<w:t>${cleanAndEscape(p.initials)}</w:t></w:r></w:p></w:tc>
      <w:tc>${VCENTER}<w:p/></w:tc><w:tc>${VCENTER}<w:p/></w:tc>
    </w:tr>`;
  });
  return xml + `</w:tbl>`;
};
