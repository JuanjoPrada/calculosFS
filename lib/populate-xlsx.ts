// Populado quirurgico de la plantilla Calculo_Deuda_y_Subasta.xlsx (congelada en
// /public/plantilla). Se edita el XML de la hoja "Informe Subasta" por MANIPULACION DE
// CADENA (regex), sin reparsear con DOMParser: asi los namespaces del worksheet
// (mc/x14ac/xr/xr2/xr3 + mc:Ignorable) se conservan BYTE A BYTE. Reparsear con
// DOMParser+XMLSerializer podia soltar los xmlns:xr2/xr3 (declarados pero solo citados en
// mc:Ignorable) y Excel entonces "reparaba" el fichero al abrirlo.
// Solo se escriben celdas de ENTRADA (nunca celdas con <f>), conservando su estilo (s=).

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
function escapeXml(s: string): string {
  return String(s).replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;")
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
  let t = String(s).trim()
  if (!t) return null
  // admite "29.801,88" (es) y "29801.88" (en)
  if (t.includes(",") && t.includes(".")) {
    t = t.lastIndexOf(",") > t.lastIndexOf(".") ? t.replace(/\./g, "").replace(",", ".") : t.replace(/,/g, "")
  } else if (t.includes(",")) {
    t = /,\d{1,2}$/.test(t) ? t.replace(/\./g, "").replace(",", ".") : t.replace(/,/g, "")
  }
  const n = Number.parseFloat(t)
  return Number.isNaN(n) ? null : n
}

// --- Editor de hoja por cadena (preserva namespaces exactos) ---
class StringSheet {
  xml: string
  constructor(xml: string) {
    this.xml = xml
  }

  private cellRe(ref: string): RegExp {
    return new RegExp(`<c r="${ref}"([^>]*?)(/>|>([\\s\\S]*?)</c>)`)
  }
  private styleOf(attrs: string): string {
    const m = attrs.match(/\s+s="\d+"/)
    return m ? m[0] : ""
  }

  private setCell(ref: string, inner: string, tAttr = "") {
    const re = this.cellRe(ref)
    const m = this.xml.match(re)
    if (m) {
      if ((m[3] || "").includes("<f")) return // nunca tocar formulas
      const style = this.styleOf(m[1])
      const cell = inner ? `<c r="${ref}"${style}${tAttr}>${inner}</c>` : `<c r="${ref}"${style}/>`
      this.xml = this.xml.replace(re, () => cell)
    } else {
      const cell = inner ? `<c r="${ref}"${tAttr}>${inner}</c>` : `<c r="${ref}"/>`
      this.insertCell(ref, cell)
    }
  }

  private insertCell(ref: string, cellXml: string) {
    const r = rowNumber(ref)
    const ci = colIndex(colLetters(ref))
    const rowRe = new RegExp(`(<row r="${r}"[^>]*>)([\\s\\S]*?)(</row>)`)
    const rm = this.xml.match(rowRe)
    if (rm) {
      const [, head, body, tail] = rm
      let pos = body.length
      const cellIter = body.matchAll(/<c r="([A-Z]+)\d+"/g)
      for (const cm of cellIter) {
        if (colIndex(cm[1]) > ci) {
          pos = cm.index ?? pos
          break
        }
      }
      const newBody = body.slice(0, pos) + cellXml + body.slice(pos)
      this.xml = this.xml.replace(rowRe, () => head + newBody + tail)
      return
    }
    // la fila no existe: crear <row> en orden dentro de <sheetData>
    const sdRe = /(<sheetData[^>]*>)([\s\S]*?)(<\/sheetData>)/
    const sm = this.xml.match(sdRe)
    if (!sm) return
    const [, head, body, tail] = sm
    let pos = body.length
    const rowIter = body.matchAll(/<row r="(\d+)"/g)
    for (const rmi of rowIter) {
      if (Number.parseInt(rmi[1], 10) > r) {
        pos = rmi.index ?? pos
        break
      }
    }
    const newRow = `<row r="${r}">${cellXml}</row>`
    const newBody = body.slice(0, pos) + newRow + body.slice(pos)
    this.xml = this.xml.replace(sdRe, () => head + newBody + tail)
  }

  setNumber(ref: string, value: number) {
    this.setCell(ref, `<v>${value}</v>`)
  }
  setString(ref: string, value: string) {
    if (!value) this.setCell(ref, "")
    else this.setCell(ref, `<is><t xml:space="preserve">${escapeXml(value)}</t></is>`, ' t="inlineStr"')
  }
  setDateISO(ref: string, iso: string) {
    const s = excelSerial(iso)
    if (s != null) this.setNumber(ref, s)
    else this.clear(ref)
  }
  clear(ref: string) {
    this.setCell(ref, "")
  }
}

/** Localiza el fichero XML de la hoja por su nombre (regex sobre workbook.xml + rels). */
async function findSheetPath(zip: JSZip, sheetName: string): Promise<string> {
  const wbXml = await zip.file("xl/workbook.xml")!.async("string")
  const shRe = new RegExp(`<sheet[^>]*name="${sheetName}"[^>]*?r:id="([^"]+)"`)
  const alt = new RegExp(`<sheet[^>]*?r:id="([^"]+)"[^>]*name="${sheetName}"`)
  const m = wbXml.match(shRe) || wbXml.match(alt)
  if (!m) throw new Error(`Hoja "${sheetName}" no encontrada en el libro`)
  const rid = m[1]
  const relsXml = await zip.file("xl/_rels/workbook.xml.rels")!.async("string")
  const rm = relsXml.match(new RegExp(`Id="${rid}"[^>]*Target="([^"]+)"`)) || relsXml.match(new RegExp(`Target="([^"]+)"[^>]*Id="${rid}"`))
  if (!rm) throw new Error("Relacion de hoja no encontrada")
  let target = rm[1]
  if (target.startsWith("/")) target = target.slice(1)
  else if (!target.startsWith("xl/")) target = "xl/" + target
  return target
}

/** Asegura calculo AUTOMATICO + recalculo completo al abrir (edicion de cadena).
 *  calcMode="auto" es clave: obliga a Excel a recalcular solo (sin que el usuario pulse F9),
 *  aunque su sesion de Excel estuviera en modo manual. */
function forceAutoCalc(wbXml: string): string {
  const m = wbXml.match(/<calcPr\b([^>]*?)\/?>/)
  if (m) {
    const calcId = m[1].match(/calcId="\d+"/)
    const cp = `<calcPr calcMode="auto" fullCalcOnLoad="1"${calcId ? " " + calcId[0] : ""}/>`
    return wbXml.slice(0, m.index) + cp + wbXml.slice((m.index ?? 0) + m[0].length)
  }
  return wbXml.replace(/<\/workbook>/, '<calcPr calcMode="auto" fullCalcOnLoad="1"/></workbook>')
}

// Version de la plantilla: subir este numero cada vez que se recongele la plantilla,
// para forzar que el navegador/CDN descarguen la nueva y no una cacheada corrupta.
const TEMPLATE_VERSION = "20260720c"

export async function populateAndDownload(data: InformeInput): Promise<void> {
  const res = await fetch(`/plantilla/Calculo_Deuda_y_Subasta.xlsx?v=${TEMPLATE_VERSION}`, { cache: "no-store" })
  if (!res.ok) throw new Error("No se pudo cargar la plantilla del servidor.")
  const zip = await JSZip.loadAsync(await res.arrayBuffer())

  const sheetPath = await findSheetPath(zip, SHEET_NAME)
  const sheet = new StringSheet(await zip.file(sheetPath)!.async("string"))

  const num = (s: string, ref: string, fallbackZero = false) => {
    const n = toNumber(s)
    if (n != null) sheet.setNumber(ref, n)
    else if (fallbackZero) sheet.setNumber(ref, 0)
    else sheet.clear(ref)
  }

  // 1 · expediente
  sheet.setString("D8", data.loanNumber)
  sheet.setString("K8", data.legalId)
  sheet.setString("R8", data.loanNumber)
  sheet.clear("Y8")

  // 2 · datos economicos
  num(data.principal, "D10", true)
  num(data.interest, "D11", true)
  num(data.penaltyADE, "D12", true)
  num(data.expenses, "D13", true)
  num(data.judicialCosts, "D14", true)
  sheet.setDateISO("D15", data.accountClosureDate)
  sheet.setDateISO("D16", data.calculationDate)
  sheet.setNumber("K14", 0)
  sheet.setNumber("K15", 0)

  // 3 · parametros de interes (tipos en FRACCION dentro del libro)
  sheet.setString("D20", data.mode || "FIJO")
  sheet.setString("D21", data.consumer || "N")
  const iar = toNumber(data.interestRatePct)
  const piar = toNumber(data.penaltyRatePct)
  if (iar != null) sheet.setNumber("D22", iar / 100)
  if (piar != null) sheet.setNumber("D23", piar / 100)
  sheet.setString("D24", "Sí")

  // 4 · parametros de subasta
  sheet.setString("R10", data.regimen12025 === "Y" ? "Y" : "N")
  num(data.budgetAmount, "R11", true)
  num(data.executionOrderAmount, "Y10", true)
  sheet.setString("R12", data.comentarios)

  // 6 · tabla por finca (filas 30-59): limpiar entradas y escribir
  for (let r = FINCA_FIRST_ROW; r <= FINCA_LAST_ROW; r++) {
    for (const col of ["A", "B", "C", "D", "I", "J", "K", "L", "Z", "AA", "AB", "AC", "AD", "AE"]) {
      sheet.clear(`${col}${r}`)
    }
  }
  data.fincas.slice(0, 30).forEach((f, i) => {
    const r = FINCA_FIRST_ROW + i
    const pid = f.propertyId.trim()
    if (/^\d+$/.test(pid)) sheet.setNumber(`A${r}`, Number(pid))
    else if (pid) sheet.setString(`A${r}`, pid)
    sheet.setString(`B${r}`, f.firstResidence === "Y" ? "Y" : "N")
    sheet.setString(`C${r}`, "Loanco")
    const base = toNumber(f.baseAmount)
    if (base != null) sheet.setNumber(`D${r}`, base)
    const msa = (s: string, col: string) => sheet.setNumber(`${col}${r}`, toNumber(s) ?? 0)
    msa(f.principalMSA, "I")
    msa(f.interestMSA, "J")
    msa(f.penaltyMSA, "K")
    msa(f.otherJudMSA, "L")
  })

  zip.file(sheetPath, sheet.xml)
  const wbXml = await zip.file("xl/workbook.xml")!.async("string")
  zip.file("xl/workbook.xml", forceAutoCalc(wbXml))

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
