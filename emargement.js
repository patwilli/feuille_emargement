/**

 * Génère les feuilles d'émargement à partir de la feuille "Informations generales".

 * Version optimisée A4 : tableau responsive avec horaires.

 */

function main(workbook: ExcelScript.Workbook) {

  const SHEET_SRC = "Informations generales";

  const PREVIEW_SHEET = "Répartition - Aperçu";

  const OUT_SHEET = "Émargement - Tous groupes";

  const TITLE = "FEUILLE D'ÉMARGEMENT";

  const TABLE_USABLE_WIDTH_CM = 26;



  const CM_TO_PT = 28.3464567;

  const LIGHT_GREY_FILL = "#F4F4F4";

  const BORDER_COLOR = "#A0A0A0";

  const TITLE_ROW_HEIGHT_PT = 20;

  const SUBTITLE_ROW_HEIGHT_PT = 12;

  const INFO_ROW_HEIGHT_PT = 14;

  const GAP_BETWEEN_INFO_AND_TABLE_PT = 14; // hauteur (pt) de l'espace entre le cartouche "Lieu" et le tableau

  const DISCLAIMER_ROW_HEIGHT_PT = 13;

  const MIN_PART_HPT_SIGNATURE = 23;

  const FOOTER_TEXT = "SCYFCO SAS\n51 rue de Miromesnil, 75008 Paris - https://scyfco.fr/ - contact@scyfco.fr - SIRET 880 115 431 00012 - APE 8559A - TVA FR00 880 115 431\nDéclaration d'activité enregistrée sous le n° 11756182475 auprès du préfet d'Ile-de-France\n(Cet enregistrement ne vaut pas agrément de l'état)";

  const FOOTER_ROW_HEIGHT_PT = 40;

  const FOOTER_COLOR = "#0066CC"; // Bleu



  const TOTAL_COLS = 15;

  const SIG_START = 1;

  const SIG_COLS = TOTAL_COLS - 1;



  const PAGE_ROWS = 27;

  const FIRST_CONTENT_ROW = 2;

  const DISCLAIMER = "En signant, je certifie avoir lu et accepté le règlement intérieur et reçu une formation à l'utilisation des EPI.";

  const FEEDBACK_CELL = "A4";

  const MAX_GROUP_CAPACITY = 15;



  type CellV = string | number | boolean |

    null | undefined;

  interface Person { nom: string; prenom: string; }

  interface Day {

    date: string;

    matin: string;

    aprem: string;

  }

  interface UsedBlock {

    range: ExcelScript.Range; values: CellV[][]; top: number; left: number;

    rows: number; cols: number;

  }

  interface BaseInfo {

    entreprise: CellV; certificat: CellV; langue: CellV; site: CellV;

  }



  function addBoxBorder(fmt: ExcelScript.RangeFormat, boldEdges: boolean): void {

    const edges = [

      ExcelScript.BorderIndex.edgeTop,

      ExcelScript.BorderIndex.edgeBottom,

      ExcelScript.BorderIndex.edgeLeft,

      ExcelScript.BorderIndex.edgeRight

    ];

    for (const edge of edges) {

      const b = fmt.getRangeBorder(edge);

      b.setStyle(ExcelScript.BorderLineStyle.continuous);

      b.setWeight(boldEdges ? ExcelScript.BorderWeight.medium : ExcelScript.BorderWeight.thin);

      b.setColor(BORDER_COLOR);

    }

  }



  function addRightBorder(fmt: ExcelScript.RangeFormat, weight: ExcelScript.BorderWeight = ExcelScript.BorderWeight.thin): void {

    const b = fmt.getRangeBorder(ExcelScript.BorderIndex.edgeRight);

    b.setStyle(ExcelScript.BorderLineStyle.continuous);

    b.setWeight(weight);

    b.setColor(BORDER_COLOR);

  }



  function setVerticalBorder(

    ws: ExcelScript.Worksheet,

    rowTop: number,

    colIndex: number,

    rows: number,

    which: "left" | "right",

    weight: ExcelScript.BorderWeight,

    color: string = BORDER_COLOR,

    style: ExcelScript.BorderLineStyle = ExcelScript.BorderLineStyle.continuous

  ) {

    const fmt = ws.getRangeByIndexes(rowTop, colIndex, rows, 1).getFormat();

    const edge = (which === "left")

      ? ExcelScript.BorderIndex.edgeLeft

      : ExcelScript.BorderIndex.edgeRight;

    const b = fmt.getRangeBorder(edge);

    b.setStyle(style);

    b.setWeight(weight);

    b.setColor(color);

  }



  function removeAllBorders(fmt: ExcelScript.RangeFormat): void {

    try {

      const borders = fmt.getBorders();

      borders.getItem(ExcelScript.BorderIndex.all).setStyle(ExcelScript.BorderLineStyle.none);

    } catch (e) {

    }



    const allBorders = [

      ExcelScript.BorderIndex.edgeTop,

      ExcelScript.BorderIndex.edgeBottom,

      ExcelScript.BorderIndex.edgeLeft,

      ExcelScript.BorderIndex.edgeRight,

      ExcelScript.BorderIndex.insideHorizontal,

      ExcelScript.BorderIndex.insideVertical

    ];



    for (const edge of allBorders) {

      const border = fmt.getRangeBorder(edge);

      border.setStyle(ExcelScript.BorderLineStyle.none);

      border.setWeight(ExcelScript.BorderWeight.hairline);

    }



    const edges = [

      ExcelScript.BorderIndex.edgeTop,

      ExcelScript.BorderIndex.edgeBottom,

      ExcelScript.BorderIndex.edgeLeft,

      ExcelScript.BorderIndex.edgeRight

    ];



    for (const edge of edges) {

      fmt.getRangeBorder(edge).setStyle(ExcelScript.BorderLineStyle.none);

    }

  }



  function addThinGridBorders(fmt: ExcelScript.RangeFormat): void {

    const borders = [

      ExcelScript.BorderIndex.edgeTop,

      ExcelScript.BorderIndex.edgeBottom,

      ExcelScript.BorderIndex.edgeLeft,

      ExcelScript.BorderIndex.edgeRight,

      ExcelScript.BorderIndex.insideHorizontal,

      ExcelScript.BorderIndex.insideVertical

    ];

    for (const border of borders) {

      const b = fmt.getRangeBorder(border);

      b.setStyle(ExcelScript.BorderLineStyle.continuous);

      b.setWeight(ExcelScript.BorderWeight.thin);

      b.setColor(BORDER_COLOR);

    }

  }



  function toTitleCase(str: string): string {

    return str.toLowerCase().split(' ').map(word =>

      word.charAt(0).toUpperCase() + word.slice(1)

    ).join(' ');

  }



  function toColLetter(n: number): string {

    let s = "";

    while (n > 0) {

      const m = (n - 1) % 26;

      s = String.fromCharCode(65 + m) + s;

      n = ((n - 1) / 26) | 0;

    }

    return s;

  }



  function fmtDate(v: CellV): string {

    if (typeof v === "number") {

      const excelEpoch = new Date(1900, 0, 1);

      const daysOffset = v - 2;

      const dateMs = excelEpoch.getTime() + (daysOffset * 86400000);

      const d = new Date(dateMs);

      const dd = String(d.getDate()).padStart(2, "0");

      const mm = String(d.getMonth() + 1).padStart(2, "0");

      const yy = d.getFullYear();

      return `${dd}/${mm}/${yy}`;

    }

    const s = String(v ?? "").trim();

    if (/^\d{2}\/\d{2}\/\d{4}$/.test(s)) return s;

    return s;

  }



  function fmtTime(v: CellV): string {

    if (typeof v === "number" && v >= 0 && v < 1) {

      const mins = Math.round(v * 1440);

      const hh = String(Math.floor(mins / 60)).padStart(2, "0");

      const mm = String(mins % 60).padStart(2, "0");

      return `${hh}:${mm}`;

    }

    const s = String(v ?? "").trim();

    const m = s.match(/^(\d{1,2}):(\d{1,2})$/);

    return m ?

      `${m[1].padStart(2, "0")}:${m[2].padStart(2, "0")}` : s;

  }



  function norm(raw: CellV): string {

    return String(raw ?? "")

      .normalize("NFD").replace(/[\u0300-\u036f]/g, "")

      .replace(/[:;.,()/_-]/g, " ").replace(/\s+/g, " ")

      .trim().toLowerCase();

  }



  function getUsedBlock(ws: ExcelScript.Worksheet): UsedBlock {

    const range = ws.getUsedRange(true);

    const values = range.getValues() as CellV[][];

    return { range, values, top: range.getRowIndex(), left: range.getColumnIndex(), rows: range.getRowCount(), cols: range.getColumnCount() };

  }



  function findHeaderLoc(used: UsedBlock, variants: string[]) {

    const targets = variants.map(v => norm(v));

    for (let r = 0; r < used.rows; r++) {

      const row = used.values[r];

      for (let c = 0; c < used.cols; c++) {

        if (targets.includes(norm(row[c]))) return { row: used.top + r, col: used.left + c };

      }

    }

    return null;

  }



  function findNomPrenomHeaders(used: UsedBlock) {

    const isNom = (t: string) => /^(nom|noms)( des?)?( participants?)?$/.test(t) && !t.includes("entreprise");

    const isPre = (t: string) => /^(prenom|prenoms)( des?)?( participants?)?$/.test(t);

    for (let r = 0; r < used.rows; r++) {

      let cN = -1, cP = -1;

      for (let c = 0; c < used.cols; c++) {

        const t = norm(used.values[r][c]);

        if (cN === -1 && isNom(t)) cN = used.left + c;

        if (cP === -1 && isPre(t)) cP = used.left + c;

        if (cN !== -1 && cP !== -1) return { row: used.top + r, colNom: cN, colPrenom: cP };

      }

    }

    return null;

  }



  function writeUnderHeader(ws: ExcelScript.Worksheet, used: UsedBlock, headerVariants: string[], value: CellV): boolean {

    const pos = findHeaderLoc(used, headerVariants);

    if (!pos) return false;



    const target = ws.getRangeByIndexes(pos.row + 1, pos.col, 1, 1);

    target.setValue(value);



    const f = target.getFormat();

    f.getFont().setBold(true);

    f.getFill().setColor("#E2EFDA");



    return true;

  }



  function readNumberUnderHeader(ws: ExcelScript.Worksheet, used: UsedBlock, headerVariants: string[], defaultVal: number): number {

    const pos = findHeaderLoc(used, headerVariants);

    if (!pos) return defaultVal;

    const v = ws.getRangeByIndexes(pos.row + 1, pos.col, 1, 1).getValue();

    const n = Number(v);

    return (isFinite(n) ? n : defaultVal);

  }



  const VAR_NB_GROUPES = [

    "Nombre de groupe", "Nombre de groupes",

    "Nb groupe", "Nb groupes", "Nb de groupes"

  ];



  const VAR_MAX_PAX = [

    "Max pax/groupe", "Capacité max/groupe", "Capacite max/groupe",

    "Max par groupe", "Taille max/groupe"

  ];



  function getOrCreateMentorsTable(ws: ExcelScript.Worksheet): { topRow: number; leftCol: number } {

    const used = getUsedBlock(ws);

    const hdr = findHeaderLoc(used, ["Mentor", "Mentors", "MENTORS"]);

    if (hdr) {

      return { topRow: hdr.row, leftCol: hdr.col };

    }



    const startRow = 12;

    const startCol = 1;

    const header = ws.getRangeByIndexes(startRow, startCol, 1, 2);

    header.setValues([["N° groupe", "MENTORS"]]);

    const hf = header.getFormat();

    hf.getFont().setBold(true);

    hf.getFill().setColor("#F2F2F2");

    addThinGridBorders(hf);



    return { topRow: startRow, leftCol: startCol + 1 };

  }



  function readExistingMentorsFromTable(ws: ExcelScript.Worksheet, topRow: number, leftColMentor: number): Record<number, string> {

    const map: Record<number, string> = {};

    const used = ws.getUsedRange(true);

    const maxRows = used.getRowIndex() + used.getRowCount();

    const colGroup = leftColMentor - 1;



    for (let r = topRow + 1; r < maxRows; r++) {

      const g = Number(String(ws.getRangeByIndexes(r, colGroup, 1, 1).getValue() ?? "").trim());

      const m = String(ws.getRangeByIndexes(r, leftColMentor, 1, 1).getValue() ?? "").trim();

      if (!g && !m) break;

      if (!isNaN(g) && g > 0) map[g] = m;

    }

    return map;

  }



  function ensureMentorsTable(ws: ExcelScript.Worksheet, groupsCount: number): Record<number, string> {

    const meta = getOrCreateMentorsTable(ws);

    const old = readExistingMentorsFromTable(ws, meta.topRow, meta.leftCol);



    const totalRows = Math.max(groupsCount + 1, 2);

    ws.getRangeByIndexes(meta.topRow, meta.leftCol - 1, totalRows, 2)

      .clear(ExcelScript.ClearApplyTo.all);



    const header = ws.getRangeByIndexes(meta.topRow, meta.leftCol - 1, 1, 2);

    header.setValues([["N° groupe", "MENTORS"]]);

    const hf = header.getFormat();

    hf.getFont().setBold(true);

    hf.getFill().setColor("#F2F2F2");

    addThinGridBorders(hf);



    const data: (string | number)[][] = [];

    for (let i = 1; i <= groupsCount; i++) {

      data.push([i, old[i] ?? ""]);

    }

    const dataRange = ws.getRangeByIndexes(meta.topRow + 1, meta.leftCol - 1, groupsCount, 2);

    dataRange.setValues(data);

    addThinGridBorders(dataRange.getFormat());



    const mentorColRange = ws.getRangeByIndexes(meta.topRow + 1, meta.leftCol, groupsCount, 1);

    mentorColRange.getFormat().getFill().setColor("#FFF6CC");



    const updated: Record<number, string> = {};

    for (let i = 1; i <= groupsCount; i++) updated[i] = old[i] ?? "";

    return updated;

  }



  function setFeedback(ws: ExcelScript.Worksheet, msg: string): void {

    const cell = ws.getRange(FEEDBACK_CELL);

    cell.setValue("⚠️ " + msg);

    const f = cell.getFormat();

    f.setWrapText(true);

    f.getFill().setColor("#FFF2CC");

    f.getFont().setColor("#9C5700");

  }



  function showStopOnSource(msg: string, ws?: ExcelScript.Worksheet, focusAddress?: string): never {

    try {

      const target = ws ??

        workbook.getWorksheet(SHEET_SRC);

      if (target) {

        setFeedback(target, msg);

        target.activate();

        if (focusAddress) target.getRange(focusAddress).select();

      }

    } catch { }

    throw new Error(msg);

  }



  function validateMentorsCount(map: Record<number, string>, expected: number, _w: string, _n: string, srcWs: ExcelScript.Worksheet, focusAddr: string): void {

    let nonEmptyCount = 0;

    const missing: number[] = [];

    for (let i = 1; i <= expected; i++) {

      const v = (map[i] ?? "").toString().trim();

      if (v) nonEmptyCount++; else missing.push(i);

    }

    const extras = Object.keys(map).map(k => parseInt(k, 10))

      .filter(k => !isNaN(k) && (k < 1 || k > expected) && (map[k] ?? "").toString().trim() !== "");

    if (nonEmptyCount !== expected || extras.length > 0) {

      let msg = `Le nombre de mentors doit être exactement égal au nombre de groupes.`;

      if (missing.length > 0) msg += `\nGroupes sans mentor : ${missing.join(", ")}.`;

      msg += `\nComplétez/ajustez la zone "N° groupe / MENTORS".`;

      showStopOnSource(msg, srcWs, focusAddr);

    }

  }



  function validateCapacityConstraintsOnPreview(previewWs: ExcelScript.Worksheet, groups: Person[][], groupCapacities: number[], limitPerGroup: number): void {

    let capTarget = Number(previewWs.getRange("B2").getValue());

    if (!isFinite(capTarget) || capTarget <= 0 || capTarget > limitPerGroup) capTarget = limitPerGroup;



    const offenders: { group: number; capacity: number; actual: number }[] = [];

    for (let i = 0; i < groups.length; i++) {

      const actualCount = groups[i].length;

      const totalCapacity = groupCapacities[i] || actualCount;

      if (totalCapacity > limitPerGroup || actualCount > limitPerGroup) {

        offenders.push({ group: i + 1, capacity: totalCapacity, actual: actualCount });

      }

    }

    if (offenders.length > 0) {

      const details = offenders.map(o =>

        (o.capacity > limitPerGroup)

          ? `Groupe ${o.group} (Total: ${o.capacity})`

          : `Groupe ${o.group} (${o.actual} participants)`

      ).join(", ");

      showStopOnSource(

        `Certain(s) groupe(s) dépassent la capacité maximale de ${limitPerGroup} : ${details}.\n` +

        `→ Répartissez différemment dans "${PREVIEW_SHEET}" (≤ ${limitPerGroup}/groupe), puis relancez.`,

        previewWs, "A3"

      );

    }

  }



  const src = workbook.getWorksheet(SHEET_SRC);

  if (!src) throw new Error(`La feuille "${SHEET_SRC}" est introuvable.`);

  try {

    src.getRange(FEEDBACK_CELL).clear(ExcelScript.ClearApplyTo.contents);

  } catch { }

  const usedSrc = getUsedBlock(src);



  function getValueRightOf(used: UsedBlock, labels: string[], fallbackAddr: string): CellV {

    const pos = findHeaderLoc(used, labels);

    if (pos) return src.getRangeByIndexes(pos.row, pos.col + 1, 1, 1).getValue();

    return src.getRange(fallbackAddr).getValue();

  }



  const baseInfo: BaseInfo = {

    entreprise: src.getRange("B5").getValue(),

    certificat: src.getRange("B6").getValue(),

    langue: src.getRange("B7").getValue(),

    site: src.getRange("B8").getValue()

  };



  const jourHdr = findHeaderLoc(usedSrc, ["Jour", "JOUR"]);

  if (!jourHdr) showStopOnSource(`En-tête "Jour" introuvable sur "${SHEET_SRC}".`, src, "A1");

  const rHdrRel = jourHdr.row - usedSrc.top;

  const cHdrRel = jourHdr.col - usedSrc.left;



  const days: Day[] = [];

  for (let rr = rHdrRel + 1; rr < usedSrc.rows; rr++) {

    const vDate = usedSrc.values[rr][cHdrRel];

    if (vDate === "" || vDate == null) break;

    const dStartM = usedSrc.values[rr][cHdrRel + 1];

    const dEndM = usedSrc.values[rr][cHdrRel + 2];

    const dStartA = usedSrc.values[rr][cHdrRel + 3];

    const dEndA = usedSrc.values[rr][cHdrRel + 4];

    const matin = (dStartM === "" || dEndM === "") ? "" : `${fmtTime(dStartM)} - ${fmtTime(dEndM)}`;

    const aprem = (dStartA === "" || dEndA === "") ? "" : `${fmtTime(dStartA)} - ${fmtTime(dEndA)}`;

    days.push({ date: fmtDate(vDate), matin, aprem });

  }

  if (days.length === 0) showStopOnSource("Aucun jour de formation trouvé.", src, "A1");



  const np = findNomPrenomHeaders(usedSrc);

  if (!np) showStopOnSource(`En-têtes "NOM" / "PRÉNOM" (participants) introuvables.`, src, "A1");

  const startRowParticipants = np.row + 1;

  const colNom = np.colNom; const colPrenom = np.colPrenom;

  const startRowRel = startRowParticipants - usedSrc.top;

  const colNomRel = colNom - usedSrc.left;

  const colPrenomRel = colPrenom - usedSrc.left;



  const participants: Person[] = [];

  for (let rr = startRowRel; rr < usedSrc.rows; rr++) {

    const n = String(usedSrc.values[rr][colNomRel] ?? "").trim();

    const p = String(usedSrc.values[rr][colPrenomRel] ?? "").trim();

    if (!n && !p) break;

    participants.push({ nom: n, prenom: p });

  }



  const maxDays = days.length;

  if (maxDays > 7) {

    showStopOnSource(`❌ La formation dure **${maxDays} jours** (au-delà de la limite de 7 jours). Annulation de la génération.`, src, "A1");

  }



  // Vérification de la longueur des noms uniquement si formation > 3 jours

  if (maxDays > 3) {

    const MAX_NAME_LENGTH = 30;

    const tooLongNames = participants.filter(p =>

      (String(p.nom).length + String(p.prenom).length) > MAX_NAME_LENGTH

    );



    if (tooLongNames.length > 0) {

      const firstLongNames = tooLongNames.slice(0, 5).map(p =>

        `${toTitleCase(p.nom)} ${toTitleCase(p.prenom)} (${String(p.nom).length + String(p.prenom).length} caractères)`

      ).join(', ');



      showStopOnSource(`❌ **${tooLongNames.length} participant(s)** ont un Nom + Prénom dépassant **${MAX_NAME_LENGTH} caractères**. Exemples: ${firstLongNames}. Veuillez corriger.`, src, "A1");

    }

  }



  let usedSrcNow = getUsedBlock(src);



  let userGroupsPref = readNumberUnderHeader(src, usedSrcNow, VAR_NB_GROUPES, 0);

  let userMaxPaxPref = readNumberUnderHeader(src, usedSrcNow, VAR_MAX_PAX, 15);



  const HARD_MAX = MAX_GROUP_CAPACITY;

  let effectiveMaxPax = Math.min(Math.max(1, Math.floor(userMaxPaxPref || 15)), HARD_MAX);



  let groupsCount = (userGroupsPref > 0)

    ? Math.floor(userGroupsPref)

    : Math.max(1, Math.ceil(participants.length / effectiveMaxPax));



  if (userGroupsPref <= 0) {

    writeUnderHeader(src, usedSrcNow, VAR_NB_GROUPES, groupsCount);

  }

  writeUnderHeader(src, usedSrcNow, VAR_MAX_PAX, effectiveMaxPax);



  usedSrcNow = getUsedBlock(src);

  userGroupsPref = readNumberUnderHeader(src, usedSrcNow, VAR_NB_GROUPES, 0);

  userMaxPaxPref = readNumberUnderHeader(src, usedSrcNow, VAR_MAX_PAX, 15);

  effectiveMaxPax = Math.min(Math.max(1, Math.floor(userMaxPaxPref || 15)), HARD_MAX);

  groupsCount = (userGroupsPref > 0)

    ? Math.floor(userGroupsPref)

    : Math.max(1, Math.ceil(participants.length / effectiveMaxPax));



  if (participants.length > groupsCount * effectiveMaxPax) {

    const minNeeded = Math.ceil(participants.length / effectiveMaxPax);

    const posG = findHeaderLoc(usedSrcNow, VAR_NB_GROUPES);

    const focusAddr = posG ? `${toColLetter(posG.col + 1)}${posG.row + 2}` : "A1";

    showStopOnSource(

      `Impossible de répartir ${participants.length} participant(s) en ${groupsCount} groupe(s) avec ${effectiveMaxPax} max/groupe.\n` +

      `→ Augmentez "Nombre de groupe" à ${minNeeded} (ou augmentez "Max pax/groupe" ≤ ${HARD_MAX}).`,

      src, focusAddr

    );

  }



  const mentorsByGroup = ensureMentorsTable(src, groupsCount);



  const mtMeta = getOrCreateMentorsTable(src);

  const firstMentorCellAddr = `${toColLetter(mtMeta.leftCol + 1)}${mtMeta.topRow + 2}`;



  validateMentorsCount(mentorsByGroup, groupsCount, "source", SHEET_SRC, src, firstMentorCellAddr);



  const defaultCapacity = Math.ceil(Math.max(1, participants.length) / groupsCount);



  let preview = workbook.getWorksheet(PREVIEW_SHEET);

  if (!preview) {

    const initialGroups: Person[][] = Array.from({ length: groupsCount }, () => [] as Person[]);

    let gi = 0;

    for (let i = 0; i < participants.length; i++) {

      initialGroups[gi].push(participants[i]);

      gi = (gi + 1) % groupsCount;

    }

    preview = workbook.addWorksheet(PREVIEW_SHEET);

    preview.getRange("A1").setValue("Répartition proposée");

    preview.getRange("A2").setValue("Capacité cible par groupe :");

    preview.getRange("B2").setValue(Math.min(defaultCapacity, effectiveMaxPax));

    preview.getRange("D2").setValue("→ Modifiez librement les colonnes ci-dessous, puis relancez la génération.");

    preview.getRange("A1:D2").getFormat().getFont().setBold(true);

    preview.getRange("A2:D2").getFormat().getFill().setColor("#FFF6CC");

    const startRowPrev = 4;

    for (let g = 0; g < initialGroups.length; g++) {

      const col = 1 + g * 2;

      const colName = toColLetter(col);

      const colFirst = toColLetter(col + 1);

      const header = preview.getRangeByIndexes(startRowPrev - 1, col - 1, 1, 2);

      header.setValues([[`Groupe ${g + 1}`, ""]]);

      header.getFormat().getFill().setColor("#F2F2F2");

      header.getFormat().getFont().setBold(true);

      const data = initialGroups[g].map(p => [p.nom, p.prenom]) as (string | number | boolean)[][];

      const dataLen = Math.max(1, data.length);

      if (data.length > 0) {

        preview.getRangeByIndexes(startRowPrev, col - 1, data.length, 2).setValues(data);

      } else {

        preview.getRangeByIndexes(startRowPrev, col - 1, 1, 2).setValues([["", ""]]);

      }



      const totalRow = startRowPrev + dataLen;

      preview.getRangeByIndexes(totalRow, col - 1, 1, 1).setValue("Total");

      preview.getRangeByIndexes(totalRow, col, 1, 1).setValue(initialGroups[g].length);

      const totalFmt = preview.getRangeByIndexes(totalRow, col - 1, 1, 2).getFormat();

      totalFmt.getFont().setBold(true);

      totalFmt.getFill().setColor("#FFF6CC");



      preview.getRange(`${colName}:${colName}`).getFormat().setColumnWidth(85);

      preview.getRange(`${colFirst}:${colFirst}`).getFormat().setColumnWidth(75);

      const borderStartRow = startRowPrev - 1;

      const rowsCovered = 1 + dataLen + 1;

      const rightColRange = preview.getRangeByIndexes(borderStartRow, col, rowsCovered, 1);

      addRightBorder(rightColRange.getFormat());

    }

    return;

  }



  const usedPrev = preview.getUsedRange(true);

  const prevVals = usedPrev.getValues() as CellV[][];

  const prevTop = usedPrev.getRowIndex();

  const prevLeft = usedPrev.getColumnIndex();

  const prevRows = usedPrev.getRowCount();

  const prevCols = usedPrev.getColumnCount();



  // FIX: Moved 'groups' declaration here to prevent TDZ error reported by user (Line 502)

  const groups: Person[][] = [];

  const groupCapacities: number[] = [];



  const headerRowAbs = 3;

  const headerRowRel = headerRowAbs - prevTop;

  if (headerRowRel < 0 || headerRowRel >= prevRows) {

    showStopOnSource(`Ligne d'en-tête introuvable dans "${PREVIEW_SHEET}".`, src, "B7");

  }





  for (let colAbs = prevLeft; colAbs < prevLeft + prevCols; colAbs += 2) {

    const colRel = colAbs - prevLeft;

    const titleText = String(prevVals[headerRowRel]?.[colRel] ?? "");

    const normalized = titleText.normalize("NFD").replace(/[\u0300-\u036f]/g, "").toLowerCase();

    if (!/^groupe\s*\d+/.test(normalized)) continue;



    const gArr: Person[] = [];

    let totalCapacity = 0;

    for (let rr = headerRowRel + 1; rr < prevRows; rr++) {

      const firstCell = String(prevVals[rr]?.[colRel] ?? "").trim();

      const isTotalRow = /^total\s*:?\s*$/i.test(firstCell);

      if (isTotalRow) {

        const capacityCell = prevVals[rr]?.[colRel + 1];

        totalCapacity = Number(capacityCell) || gArr.length;

        break;

      }

      const nom = firstCell;

      const prenom = String(prevVals[rr]?.[colRel + 1] ?? "").trim();

      if (!nom && !prenom) continue;

      gArr.push({ nom, prenom });

    }

    if (gArr.length) {

      groups.push(gArr);

      groupCapacities.push(totalCapacity || gArr.length);

    }

  }

  if (groups.length === 0) showStopOnSource(`Aucun groupe lisible dans "${PREVIEW_SHEET}".`, src, "B7");

  validateMentorsCount(mentorsByGroup, groups.length, "preview", SHEET_SRC, src, "B7");

  validateCapacityConstraintsOnPreview(preview, groups, groupCapacities, effectiveMaxPax);



  const sheets = workbook.getWorksheets();

  for (let i = sheets.length - 1; i >= 0; i--) {

    const ws = sheets[i];

    const n = ws.getName().toLowerCase();

    if (n.startsWith("émargement - groupe") || n.startsWith("emargement - groupe") || n === OUT_SHEET.toLowerCase()) {

      ws.delete();

    }

  }



  const out = workbook.addWorksheet(OUT_SHEET);



  const layout = out.getPageLayout();

  layout.setOrientation(ExcelScript.PageOrientation.landscape);

  layout.setPaperSize(ExcelScript.PaperType.a4);

  layout.setTopMargin(10);

  layout.setBottomMargin(10);

  layout.setLeftMargin(10);

  layout.setRightMargin(10);

  layout.setHeaderMargin(0);

  layout.setFooterMargin(0);

  const lo = layout as ExcelScript.PageLayout & {

    setFitToPagesWide?: (n: number) => void;

    setFitToPagesTall?: (n: number) => void;

    setCenterHorizontally?: (b: boolean) => void;

  };

  if (typeof lo.setFitToPagesWide === "function") lo.setFitToPagesWide(1);

  if (typeof lo.setCenterHorizontally === "function") lo.setCenterHorizontally(true);



  let currentTop = FIRST_CONTENT_ROW;

  for (let i = 0; i < groups.length; i++) {

    if (i > 0) addHardPageBreak(out, currentTop);

    const groupStart = currentTop;

    const afterBlockRow = writeEmargementBlockResponsive(

      out,

      groupStart,

      groups[i],

      days,

      baseInfo,

      mentorsByGroup,

      i + 1,

      groups.length

    );

    const usedRows = afterBlockRow - groupStart;

    if (usedRows < PAGE_ROWS) {

      const filler = PAGE_ROWS - usedRows;



      const refH = MIN_PART_HPT_SIGNATURE;



      const activeColCount = 1 + countActiveSignatureColumns(days);



      const fillerRange = out.getRangeByIndexes(afterBlockRow - 1, 0, filler, activeColCount);

      fillerRange.setValues(Array.from({ length: filler }, () => Array(activeColCount).fill("")));

      const ffmt = fillerRange.getFormat();

      ffmt.setRowHeight(refH);

    }



    // Ajouter le footer SCYFCO en position fixe (ligne 26 de chaque page)

    const footerRowAbs = groupStart + PAGE_ROWS - 1; // Ligne 26 (0-based = 25)

    const footerRange = out.getRangeByIndexes(footerRowAbs, 0, 1, TOTAL_COLS);

    footerRange.merge(false);

    footerRange.setValue(FOOTER_TEXT);



    const footerFmt = footerRange.getFormat();

    footerFmt.getFont().setSize(7);

    footerFmt.getFont().setColor(FOOTER_COLOR);

    footerFmt.getFont().setBold(false);

    footerFmt.setHorizontalAlignment(ExcelScript.HorizontalAlignment.center);

    footerFmt.setVerticalAlignment(ExcelScript.VerticalAlignment.center);

    footerFmt.setWrapText(true);

    footerFmt.setRowHeight(FOOTER_ROW_HEIGHT_PT);

    removeAllBorders(footerFmt);



    currentTop = groupStart + PAGE_ROWS;

  }



  const usedAll = out.getUsedRange(true);

  layout.setPrintArea(usedAll.getAddress());



  const startCleanRow = currentTop;

  const totalColumns = TOTAL_COLS;

  const endCleanRow = 500;



  if (startCleanRow < endCleanRow) {

    const rowsToClean = endCleanRow - startCleanRow;



    const cleanUpRange = out.getRangeByIndexes(startCleanRow, 0, rowsToClean, totalColumns);



    removeAllBorders(cleanUpRange.getFormat());



    cleanUpRange.clear(ExcelScript.ClearApplyTo.formats);

  }



  out.activate();



  function addHardPageBreak(ws: ExcelScript.Worksheet, row1Based: number): void {

    const rng = ws.getRangeByIndexes(row1Based - 1, 0, 1, 1);

    const wsWithPB = ws as ExcelScript.Worksheet & { addPageBreak?: (...args: unknown[]) => void };

    try {

      wsWithPB.addPageBreak?.("Horizontal", rng);

    }

    catch {

      try {

        wsWithPB.addPageBreak?.(1, rng);

      } catch { }

    }

  }



  function countActiveSignatureColumns(daysArr: Day[]): number {

    let sessions = 0;

    for (const d of daysArr) {

      const hasM = (d.matin || "").trim() !== "" && d.matin !== "-";

      const hasA = (d.aprem || "").trim() !== "" && d.aprem !== "-";



      if (hasM) sessions++;  // +1 si matin présent

      if (hasA) sessions++;  // +1 si après-midi présent

    }

    return Math.min(Math.max(sessions, 0), SIG_COLS);

  }



  function setDynamicColumnWidths(ws: ExcelScript.Worksheet, activeSigCols: number): {

    partWpt: number;

    sigWpt: number

  } {

    const totalWpt = TABLE_USABLE_WIDTH_CM * CM_TO_PT;

    const weightParticipant = 1.6;

    const weightSignature = 1.0;

    const totalActiveWeight = weightParticipant + activeSigCols * weightSignature;

    let partWpt = (weightParticipant / totalActiveWeight) * totalWpt;

    let sigWpt = (weightSignature / totalActiveWeight) * totalWpt;



    const MIN_PART_PT = 100;

    const MIN_SIG_PT = 45;

    const MIN_UNUSED_PT = 0;

    if (partWpt < MIN_PART_PT) partWpt = MIN_PART_PT;

    if (sigWpt < MIN_SIG_PT) sigWpt = MIN_SIG_PT;

    const currentActiveTotal = partWpt + activeSigCols * sigWpt;

    if (currentActiveTotal > totalWpt) {

      const scale = totalWpt / currentActiveTotal;

      partWpt *= scale;

      sigWpt *= scale;

    }



    ws.getRange("A:A").getFormat().setColumnWidth(partWpt);

    for (let k = 0; k < SIG_COLS; k++) {

      const letter = toColLetter(2 + k);

      const width = (k < activeSigCols) ? sigWpt : MIN_UNUSED_PT;

      ws.getRange(`${letter}:${letter}`).getFormat().setColumnWidth(width);

    }



    return { partWpt: partWpt, sigWpt: sigWpt };

  }



  function writeEmargementBlockResponsive(

    ws: ExcelScript.Worksheet,

    topRowAbs: number,

    groupPeople: Person[],

    daysArr: Day[],

    base: BaseInfo,

    mentors: Record<number, string>,

    groupIndex: number,

    groupsTotal: number

  ): number {

    let currentRow = topRowAbs;



    const titleRange = ws.getRangeByIndexes(currentRow, 0, 1, TOTAL_COLS);

    titleRange.merge(false);

    titleRange.setValue(groupsTotal === 1 ? TITLE : `${TITLE} — GROUPE ${groupIndex}`.toUpperCase());



    titleRange.getFormat().setRowHeight(TITLE_ROW_HEIGHT_PT);



    let titleFmt = titleRange.getFormat();

    titleFmt.getFont().setBold(true);

    titleFmt.getFont().setSize(14);

    titleFmt.getFont().setColor("#000000");

    titleFmt.getFill().clear();

    titleFmt.setHorizontalAlignment(ExcelScript.HorizontalAlignment.center);

    titleFmt.setVerticalAlignment(ExcelScript.VerticalAlignment.center);

    removeAllBorders(titleFmt);



    currentRow++;

    const dateRangeText = `${daysArr[0].date} au ${daysArr[daysArr.length - 1].date}`;

    const subtitleRange = ws.getRangeByIndexes(currentRow, 0, 1, TOTAL_COLS);

    subtitleRange.merge(false);

    subtitleRange.setValue(`(du ${dateRangeText})`);



    subtitleRange.getFormat().setRowHeight(SUBTITLE_ROW_HEIGHT_PT);

    let subtitleFmt = subtitleRange.getFormat();

    subtitleFmt.getFont().setSize(10);

    subtitleFmt.setHorizontalAlignment(ExcelScript.HorizontalAlignment.center);

    subtitleFmt.setVerticalAlignment(ExcelScript.VerticalAlignment.center);

    removeAllBorders(subtitleFmt);



    currentRow++;

    let r = currentRow;

    const mentorName = String(mentors[groupIndex] ?? "").trim();

    // Ne plus afficher le mentor dans le cartouche (il sera inclus dans la liste des participants)

    const cartoucheVals: CellV[] = [base.entreprise, base.certificat, base.site];

    const cartoucheLabels: string[] = ["Bénéficiaire", "Formation", "Lieu"];



    for (let i = 0; i < cartoucheLabels.length; i++) {

      const rowAbs = r - 1 + i;

      const label = cartoucheLabels[i];

      const value = cartoucheVals[i] ? String(cartoucheVals[i]) : "(non renseigné)";

      const infoRange = ws.getRangeByIndexes(rowAbs, 0, 1, TOTAL_COLS);

      infoRange.merge(false);

      infoRange.setValue(`${label} : ${value}`);



      const fmt = infoRange.getFormat();

      fmt.setHorizontalAlignment(ExcelScript.HorizontalAlignment.center);

      fmt.setVerticalAlignment(ExcelScript.VerticalAlignment.center);

      fmt.getFont().setSize(10);

      fmt.getFont().setColor("#333333");

      fmt.setRowHeight(INFO_ROW_HEIGHT_PT);

      removeAllBorders(fmt);

    }

    r += cartoucheLabels.length;



    // NEW — insérer une ligne vide fusionnée entre "Lieu" et le tableau

    const gap = ws.getRangeByIndexes(r - 1, 0, 1, TOTAL_COLS); // <-- r - 1 (0-based)

    gap.merge(false);

    gap.setValue("");

    const gapFmt = gap.getFormat();

    gapFmt.setRowHeight(GAP_BETWEEN_INFO_AND_TABLE_PT);

    gapFmt.getFill().clear();

    removeAllBorders(gapFmt);

    r += 1;

    // --- fin NEW



    // (optionnel) garde cette variable si tu en as besoin plus bas

    // const firstRowOfTable = r;



    // Préparer la liste : si un mentor est renseigné, l'insérer en première position

    const peopleWithMentor: Person[] = groupPeople.slice();

    if (mentorName) {

      // On met le mentor en première ligne sous la forme "Mentor. <nom-complet>"

      peopleWithMentor.unshift({ nom: `Mentor: ${mentorName}`, prenom: "" });

    }



    const activeSigCols = Math.max(1, countActiveSignatureColumns(daysArr));

    setDynamicColumnWidths(ws, activeSigCols);

    r = writeTableOfDaysResponsive(ws, r, peopleWithMentor, daysArr, activeSigCols);



    const spacerBeforeDisclaimer = ws.getRangeByIndexes(r - 1, 0, 1, TOTAL_COLS);

    spacerBeforeDisclaimer.merge(false);

    spacerBeforeDisclaimer.getFormat().setRowHeight(4);



    spacerBeforeDisclaimer.getFormat().getRangeBorder(ExcelScript.BorderIndex.edgeTop).setStyle(ExcelScript.BorderLineStyle.continuous);

    spacerBeforeDisclaimer.getFormat().getRangeBorder(ExcelScript.BorderIndex.edgeTop).setWeight(ExcelScript.BorderWeight.hairline);

    spacerBeforeDisclaimer.getFormat().getRangeBorder(ExcelScript.BorderIndex.edgeTop).setColor(BORDER_COLOR);



    r += 1;



    const disclaimerRange = ws.getRangeByIndexes(r - 1, 0, 1, TOTAL_COLS);

    disclaimerRange.merge(false);

    disclaimerRange.setValue(`* ${DISCLAIMER}`);



    const disclaimerFmt = disclaimerRange.getFormat();

    disclaimerFmt.getFont().setItalic(true);

    disclaimerFmt.getFont().setBold(true);

    disclaimerFmt.getFont().setSize(11);

    disclaimerFmt.getFont().setColor("#666666");

    disclaimerFmt.setHorizontalAlignment(ExcelScript.HorizontalAlignment.left); // ← Remis à gauche

    disclaimerFmt.setVerticalAlignment(ExcelScript.VerticalAlignment.center);

    disclaimerRange.getFormat().setRowHeight(DISCLAIMER_ROW_HEIGHT_PT);

    r += 1;



    ws.getRangeByIndexes(r - 1, TOTAL_COLS - 1, 1, 1).setValue("");



    return r;

  }



  function writeTableOfDaysResponsive(

    ws: ExcelScript.Worksheet,

    topRowAbs: number,

    groupPeople: Person[],

    daysSlice: Day[],

    activeSigCols: number

  ): number {

    const nbJ = daysSlice.length;

    const headRow0 = topRowAbs - 1;



    const headParticipant = ws.getRangeByIndexes(headRow0, 0, 2, 1);

    headParticipant.merge(false);

    headParticipant.setValue("NOM PRENOM");

    const headFixedFmt = headParticipant.getFormat();

    headFixedFmt.getFill().setColor("#F2F2F2");

    headFixedFmt.getFont().setBold(true);

    headFixedFmt.setHorizontalAlignment(ExcelScript.HorizontalAlignment.center);

    headFixedFmt.setVerticalAlignment(ExcelScript.VerticalAlignment.center);

    addThinGridBorders(headFixedFmt);



    let currentCol = SIG_START;

    let sessionIndex = 0;

    const dayEnds: number[] = [];

    const dayBlocks: { start: number; end: number }[] = [];



    for (let d = 0; d < nbJ && sessionIndex < activeSigCols; d++) {

      const hasMatin = (daysSlice[d].matin || "").trim() !== "" && daysSlice[d].matin !== "-";

      const hasAprem = (daysSlice[d].aprem || "").trim() !== "" && daysSlice[d].aprem !== "-";



      if (!hasMatin && !hasAprem) continue;  // Ignorer jour sans session



      let sessionsThisDay = 0;

      const blockStart = currentCol;



      // Ajouter matin si présent

      if (hasMatin && sessionIndex < activeSigCols) {

        sessionsThisDay++;

        sessionIndex++;

      }



      // Ajouter après-midi si présent

      if (hasAprem && sessionIndex < activeSigCols) {

        sessionsThisDay++;

        sessionIndex++;

      }



      if (sessionsThisDay === 0) continue;



      const rgDate = ws.getRangeByIndexes(headRow0, currentCol, 1, sessionsThisDay);

      rgDate.merge(false);

      rgDate.setNumberFormatLocal("@");

      rgDate.setValue(String(daysSlice[d].date));

      const f1 = rgDate.getFormat();

      f1.getFill().setColor("#F2F2F2");

      f1.getFont().setBold(true);

      f1.setHorizontalAlignment(ExcelScript.HorizontalAlignment.center);

      f1.setVerticalAlignment(ExcelScript.VerticalAlignment.center);



      const row1 = headRow0 + 1;

      let colOffset = 0;



      if (hasMatin) {

        ws.getRangeByIndexes(row1, currentCol + colOffset, 1, 1).setValue(daysSlice[d].matin);

        colOffset++;

      }



      if (hasAprem) {

        ws.getRangeByIndexes(row1, currentCol + colOffset, 1, 1).setValue(daysSlice[d].aprem);

      }



      const rgPeriode = ws.getRangeByIndexes(row1, currentCol, 1, sessionsThisDay);

      const f2 = rgPeriode.getFormat();

      f2.getFill().setColor("#EDF2F7");

      f2.getFont().setBold(true);

      f2.getFont().setSize(8);

      f2.setHorizontalAlignment(ExcelScript.HorizontalAlignment.center);

      f2.setVerticalAlignment(ExcelScript.VerticalAlignment.center);



      dayBlocks.push({ start: blockStart, end: currentCol + sessionsThisDay - 1 });

      dayEnds.push(currentCol + sessionsThisDay);

      currentCol += sessionsThisDay;

    }



    const headerSigRange = ws.getRangeByIndexes(headRow0, SIG_START, 2, activeSigCols);

    addThinGridBorders(headerSigRange.getFormat());



    for (const c of dayEnds) {

      if (c < SIG_START + activeSigCols) {

        addRightBorder(ws.getRangeByIndexes(headRow0, c - 1, 2, 1).getFormat(), ExcelScript.BorderWeight.thin);

      }

    }



    let r = topRowAbs + 2;

    let participantNumber = 1;



    for (let i = 0; i < groupPeople.length; i++) {

      const p = groupPeople[i];

      const fullName = toTitleCase(`${p.nom}${p.prenom ? ' ' + p.prenom : ''}`);



      // Détecter si la ligne est le mentor (préfixe 'Mentor.'), ne pas numéroter le mentor

      const isMentor = /^mentor[\:\s]/i.test(String(p.nom));

      const nameCellContent = isMentor ? fullName : `${participantNumber}. ${fullName}`;



      ws.getRangeByIndexes(r - 1, 0, 1, 1).setValue(nameCellContent);



      for (let c = SIG_START; c < SIG_START + activeSigCols; c++) {

        ws.getRangeByIndexes(r - 1, c, 1, 1).setValue("");

      }



      const participantActiveRange = ws.getRangeByIndexes(r - 1, 0, 1, 1 + activeSigCols);

      const dataFmt = participantActiveRange.getFormat();

      addThinGridBorders(dataFmt);



      dataFmt.setRowHeight(MIN_PART_HPT_SIGNATURE);



      const nameFmt = ws.getRangeByIndexes(r - 1, 0, 1, 1).getFormat();

      nameFmt.setHorizontalAlignment(ExcelScript.HorizontalAlignment.left);

      nameFmt.setVerticalAlignment(ExcelScript.VerticalAlignment.center);

      nameFmt.setIndentLevel(1);

      // Appliquer le gras uniquement si c'est la ligne du mentor

      nameFmt.getFont().setBold(isMentor);



      const fullNameLength = (p.nom ? String(p.nom).length : 0) + (p.prenom ? String(p.prenom).length : 0);

      if (fullNameLength > 10) {

        nameFmt.setWrapText(true);

      } else {

        nameFmt.setWrapText(false);

      }



      if (!isMentor && (participantNumber % 2 === 0)) {

        // couleur alternée basée maintenant sur la position effective des participants

        dataFmt.getFill().setColor(LIGHT_GREY_FILL);

      } else if (isMentor) {

        // Optionnel : mettre un fond spécifique pour le mentor — ici on laisse sans remplissage

        dataFmt.getFill().clear();

      } else {

        dataFmt.getFill().clear();

      }



      for (const c of dayEnds) {

        if (c < SIG_START + activeSigCols) {

          addRightBorder(ws.getRangeByIndexes(r - 1, c - 1, 1, 1).getFormat(), ExcelScript.BorderWeight.thin);

        }

      }



      if (!isMentor) participantNumber++;

      r++;

    }



    const tableTop = headRow0;

    const tableHeight = (r - headRow0);



    for (const b of dayBlocks) {

      setVerticalBorder(ws, tableTop, b.start, tableHeight, "left", ExcelScript.BorderWeight.medium);



      setVerticalBorder(ws, tableTop, b.end, tableHeight, "right", ExcelScript.BorderWeight.medium);



      if (b.end - b.start === 1) {

        setVerticalBorder(

          ws, tableTop, b.start, tableHeight,

          "right",

          ExcelScript.BorderWeight.thin,

          "#C0C0C0",

          ExcelScript.BorderLineStyle.dash

        );

      }

    }



    // Nettoyer les bordures fantômes sur les colonnes non-actives
    const firstUnusedCol = SIG_START + activeSigCols;
    const unusedColCount = TOTAL_COLS - firstUnusedCol;
    if (unusedColCount > 0) {
      const unusedRange = ws.getRangeByIndexes(
        headRow0, firstUnusedCol,
        r - headRow0, unusedColCount
      );
      removeAllBorders(unusedRange.getFormat());
      unusedRange.getFormat().getFill().clear();
    }

    // Ré-asserter la bordure droite de fermeture du tableau (peut être écrasée par le nettoyage ci-dessus)
    const lastActiveCol = SIG_START + activeSigCols - 1;
    setVerticalBorder(ws, tableTop, lastActiveCol, tableHeight, "right", ExcelScript.BorderWeight.medium);



    return r;

  }

}



