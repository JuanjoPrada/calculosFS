// Populado quirurgico de la plantilla Calculo_Deuda_y_Subasta.xlsx (congelada en
// /public/plantilla). Se edita el XML interno del .xlsx con JSZip + DOMParser: SOLO se
// escriben valores en celdas de ENTRADA de la hoja "Informe Subasta" (nunca celdas con
// formula), conservando el estilo de cada celda. Asi el libro conserva al 100% sus
// formulas, formatos, graficos y demas hojas, y recalcula al abrirse (fullCalcOnLoad).

import JSZip from "jszip"

export interface FincaInput {
  propertyId: string
  firstResidence: string // "Y" | "N" | ""
  baseAmount: string // tipo de subasta (BaseAmount); puede quedar vacio
  principalMSA: string
  interestMSA: string
  penaltyMSA: string
  otherJudMSA: string // OtherExpensesMSA + JudicialCostsMSA (columna L)
}

export interface InformeInput {
  loanNumber: string
  legalId: string
  comentarios: string
  principal: string
  interest: string
  penaltyADE: string // demora a cierre (D12)
  expenses: string
  judicialCosts: string
  accountClosureDate: string // ISO yyyy-mm-dd
  calculationDate: string
  mode: string // FIJO | VARIABLE | LEGAL
  consumer: string // Y | N
  interestRatePct: string // en % (7.5)
  penaltyRatePct: string
  regimen12025: string // Y | N
  executionOrderAmount: string
  budgetAmount: string
  fincas: FincaInput[]
}

const SHEET_NAME = "Informe Subasta"
const FINCA_FIRST_ROW = 30
const FINCA_LAST_ROW = 59

function colLetters(ref: string): string {
  return ref.replace(/\d+/g, "")
}
function rowNumber(ref: string): number {
  return Number.parseInt(ref.replace(/[A-Z]+/g, ""), 10)
}
function colIndex(letters: string): number {
  let n = 0
  for (const ch of letters) n = n * 26 + (ch.charCodeAt(0) - 64)
  return n
}

/** Serial Excel (sistema 1900) para una fecha ISO. */
function excelSerial(iso: string): number | null {
  const m = iso.match(/^(\d{4})-(\d{2})-(\d{2})$/)
  if (!m) return null
  const ms = Date.UTC(+m[1], +m[2] - 1, +m[3]) - Date.UTC(1899, 11, 30)
  return Math.round(ms / 86400000)
}

function toNumber(s: string): number | null {
  if (s == null) return null
  const t = String(s).trim().replace(/\./g, (d, i, str) => (str.indexOf(",") > -1 ? "" : d)).replace(",", ".")
  if (!t) return null
  const n = Number.parseFloat(t)
  return Number.isNaN(n) ? null : n
}

class SheetEditor {
  doc: Document
  ns: string
  rowsEl: Element
  constructor(xml: string) {
    this.doc = new DOMParser().parseFromString(xml, "application/xml")
    this.ns = this.doc.documentElement.namespaceURI || ""
    const sd = this.doc.getElementsByTagNameNS(this.ns, "sheetData")[0]
    if (!sd) throw new Error("sheetData no encontrado en la hoja")
    this.rowsEl = sd
  }

  private findRow(r: number): Element | null {
    const rows = this.rowsEl.getElementsByTagNameNS(this.ns, "row")
    for (let i = 0; i < rows.length; i++) {
      if (Number(rows[i].getAttribute("r")) === r) return rows[i]
    }
    return null
  }

  private ensureRow(r: number): Element {
    const existing = this.findRow(r)
    if (existing) return existing
    const row = this.doc.createElementNS(this.ns, "row")
    row.setAttribute("r", String(r))
    // insertar en orden
    const rows = this.rowsEl.getElementsByTagNameNS(this.ns, "row")
    let before: Element | null = null
    for (let i = 0; i < rows.length; i++) {
      if (Number(rows[i].getAttribute("r")) > r) { before = rows[i]; break }
    }
    this.rowsEl.insertBefore(row, before)
    return row
  }

  private ensureCell(ref: string): Element {
    const r = rowNumber(ref)
    const row = this.ensureRow(r)
    const cells = row.getElementsByTagNameNS(this.ns, "c")
    const myCol = colIndex(colLetters(ref))
    for (let i = 0; i < cells.length; i++) {
      if (cells[i].getAttribute("r") === ref) return cells[i]
    }
    const cell = this.doc.createElementNS(this.ns, "c")
    cell.setAttribute("r", ref)
    let before: Element | null = null
    for (let i = 0; i < cells.length; i++) {
      const c = cells[i].getAttribute("r") || ""
      if (colIndex(colLetters(c)) > myCol) { before = cells[i]; break }
    }
    row.insertBefore(cell, before)
    return cell
  }

  /** True si la celda contiene una formula (no debe tocarse). */
  hasFormula(ref: string): boolean {
    const row = this.findRow(rowNumber(ref))
    if (!row) return false
    const cells = row.getElementsByTagNameNS(this.ns, "c")
    for (let i = 0; i < cells.length; i++) {
      if (cells[i].getAttribute("r") === ref) {
        return cells[i].getElementsByTagNameNS(this.ns, "f").length > 0
      }
    }
    return false
  }

  private resetCell(cell: Element) {
    while (cell.firstChild) cell.removeChild(cell.firstChild)
    cell.removeAttribute("t")
  }

  clear(ref: string) {
    if (this.hasFormula(ref)) return
    const row = this.findRow(rowNumber(ref))
    if (!row) return
    const cells = row.getElementsByTagNameNS(this.ns, "c")
    for (let i = 0; i < cells.length; i++) {
      if (cells[i].getAttribute("r") === ref) { this.resetCell(cells[i]); return }
    }
  }

  setNumber(ref: string, value: number) {
    if (this.hasFormula(ref)) return
    const cell = this.ensureCell(ref)
    this.resetCell(cell)
    const v = this.doc.createElementNS(this.ns, "v")
    v.textContent = String(value)
    cell.appendChild(v)
  }

  setString(ref: string, value: string) {
    if (this.hasFormula(ref)) return
    const cell = this.ensureCell(ref)
    this.resetCell(cell)
    if (!value) return
    cell.setAttribute("t", "inlineStr")
    const is = this.doc.createElementNS(this.ns, "is")
    const t = this.doc.createElementNS(this.ns, "t")
    t.setAttribute("xml:space", "preserve")
    t.textContent = value
    is.appendChild(t)
    cell.appendChild(is)
  }

  setDateISO(ref: string, iso: string) {
    const serial = excelSerial(iso)
    if (serial != null) this.setNumber(ref, serial)
    else this.clear(ref)
  }

  serialize(): string {
    return new XMLSerializer().serializeToString(this.doc)
  }
}

/** Localiza el fichero XML de la hoja por su nombre en workbook.xml + rels. */
async function findSheetPath(zip: JSZip, sheetName: string): Promise<string> {
  const wbXml = await zip.file("xl/workbook.xml")!.async("string")
  const wb = new DOMParser().parseFromString(wbXml, "application/xml")
  const ns = wb.documentElement.namespaceURI || ""
  const sheets = wb.getElementsByTagNameNS(ns, "sheet")
  let rid = ""
  for (let i = 0; i < sheets.length; i++) {
    if (sheets[i].getAttribute("name") === sheetName) {
      rid = sheets[i].getAttributeNS("http://schemas.openxmlformats.org/officeDocument/2006/relationships", "id") || sheets[i].getAttribute("r:id") || ""
      break
    }
  }
  if (!rid) throw new Error(`Hoja "${sheetName}" no encontrada en el libro`)
  const relsXml = await zip.file("xl/_rels/workbook.xml.rels")!.async("string")
  const rels = new DOMParser().parseFromString(relsXml, "application/xml")
  const relEls = rels.getElementsByTagName("Relationship")
  for (let i = 0; i < relEls.length; i++) {
    if (relEls[i].getAttribute("Id") === rid) {
      let target = relEls[i].getAttribute("Target") || ""
      if (target.startsWith("/")) target = target.slice(1)
      else if (!target.startsWith("xl/")) target = "xl/" + target
      return target
    }
  }
  throw new Error("Relacion de hoja no encontrada")
}

/** Fuerza el recalculo completo al abrir el libro. */
function forceFullCalc(wbXml: string): string {
  if (/fullCalcOnLoad=/.test(wbXml)) return wbXml
  if (/<calcPr\b/.test(wbXml)) return wbXml.replace(/<calcPr\b/, '<calcPr fullCalcOnLoad="1" ')
  return wbXml.replace(/<\/workbook>/, '<calcPr fullCalcOnLoad="1"/></workbook>')
}

export async function populateAndDownload(data: InformeInput): Promise<void> {
  const res = await fetch("/plantilla/Calculo_Deuda_y_Subasta.xlsx")
  if (!res.ok) throw new Error("No se pudo cargar la plantilla del servidor.")
  const zip = await JSZip.loadAsync(await res.arrayBuffer())

  const sheetPath = await findSheetPath(zip, SHEET_NAME)
  const editor = new SheetEditor(await zip.file(sheetPath)!.async("string"))

  const num = (s: string, ref: string, fallbackZero = false) => {
    const n = toNumber(s)
    if (n != null) editor.setNumber(ref, n)
    else if (fallbackZero) editor.setNumber(ref, 0)
    else editor.clear(ref)
  }

  // --- 1 · expediente (fila 8: LoanID D8, LegalID K8, LoanNumber R8, OriginalLoanNumber Y8)
  editor.setString("D8", data.loanNumber)
  if (data.legalId) editor.setString("K8", data.legalId)
  else editor.clear("K8")
  editor.setString("R8", data.loanNumber)
  editor.clear("Y8")

  // --- 2 · datos economicos
  num(data.principal, "D10", true)
  num(data.interest, "D11", true)
  num(data.penaltyADE, "D12", true)
  num(data.expenses, "D13", true)
  num(data.judicialCosts, "D14", true)
  editor.setDateISO("D15", data.accountClosureDate)
  editor.setDateISO("D16", data.calculationDate)
  editor.setNumber("K14", 0) // SecondLienAmount: no viene de esta pantalla
  editor.setNumber("K15", 0) // ThirdLienAmount

  // --- 3 · parametros de interes (tipos en FRACCION dentro del libro)
  editor.setString("D20", data.mode || "FIJO")
  editor.setString("D21", data.consumer || "N")
  const iar = toNumber(data.interestRatePct)
  const piar = toNumber(data.penaltyRatePct)
  if (iar != null) editor.setNumber("D22", iar / 100)
  if (piar != null) editor.setNumber("D23", piar / 100)
  editor.setString("D24", "Sí")

  // --- 4 · parametros de subasta
  editor.setString("R10", data.regimen12025 === "Y" ? "Y" : "N")
  num(data.budgetAmount, "R11", true)
  num(data.executionOrderAmount, "Y10", true)
  editor.setString("R12", data.comentarios)

  // --- 6 · tabla por finca (filas 30-59): limpiar entradas y escribir
  for (let r = FINCA_FIRST_ROW; r <= FINCA_LAST_ROW; r++) {
    for (const col of ["A", "B", "C", "D", "I", "J", "K", "L", "Z", "AA", "AB", "AC", "AD", "AE"]) {
      editor.clear(`${col}${r}`)
    }
  }
  data.fincas.slice(0, 30).forEach((f, i) => {
    const r = FINCA_FIRST_ROW + i
    if (/^\d+$/.test(f.propertyId.trim())) editor.setNumber(`A${r}`, Number(f.propertyId.trim()))
    else editor.setString(`A${r}`, f.propertyId.trim())
    editor.setString(`B${r}`, f.firstResidence === "Y" ? "Y" : "N")
    editor.setString(`C${r}`, "Loanco")
    const base = toNumber(f.baseAmount)
    if (base != null) editor.setNumber(`D${r}`, base)
    const msa = (s: string, col: string) => {
      const n = toNumber(s)
      if (n != null) editor.setNumber(`${col}${r}`, n)
      else editor.setNumber(`${col}${r}`, 0)
    }
    msa(f.principalMSA, "I")
    msa(f.interestMSA, "J")
    msa(f.penaltyMSA, "K")
    msa(f.otherJudMSA, "L")
  })

  zip.file(sheetPath, editor.serialize())
  const wbXml = await zip.file("xl/workbook.xml")!.async("string")
  zip.file("xl/workbook.xml", forceFullCalc(wbXml))

  const blob = await zip.generateAsync({
    type: "blob",
    mimeType: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    compression: "DEFLATE",
  })
  const name = (data.loanNumber || data.legalId || "expediente").replace(/[^\w.-]+/g, "_")
  const url = URL.createObjectURL(blob)
  const a = document.createElement("a")
  a.href = url
  a.download = `Calculo_Deuda_y_Subasta_${name}.xlsx`
  document.body.appendChild(a)
  a.click()
  setTimeout(() => {
    document.body.removeChild(a)
    URL.revokeObjectURL(url)
  }, 0)
}
