// Parser del texto copiado (Ctrl+A, Ctrl+C) de la pantalla LegalUnpaidPrincipalBalance
// de 4Sight. Extrae por ETIQUETAS (tolerante al orden y a saltos de linea) y deja que la
// vista previa editable corrija cualquier hueco. Numeros en formato ingles (29,801.88)
// o espanol (29.801,88) segun la configuracion del navegador del usuario.

export interface PropertyRow {
  propertyId: string
  firstResidence: "Y" | "N" | ""
  principalMSA: number | null
  interestMSA: number | null
  penaltyMSA: number | null
  judicialMSA: number | null
  otherMSA: number | null
}

export interface Parsed4Sight {
  loanNumber: string
  legalId: string
  folder: string
  borrower: string
  // Debt Calculation
  accountClosureDate: string // ISO yyyy-mm-dd
  calculationDate: string
  interestRatePct: number | null // en % (7.5)
  penaltyRatePct: number | null
  consumer: "Y" | "N" | ""
  firstResidence: "Y" | "N" | ""
  updatedInterest: number | null
  judicialCosts: number | null
  // Total Updated Debt
  capital: number | null
  interest: number | null
  updatedPenaltyInterest: number | null
  otherExpenses: number | null
  totalUpdatedDebt: number | null
  // Derivado (criterio estandar): demora a cierre = UPI - UI
  penaltyADE: number | null
  properties: PropertyRow[]
  iconFieldsUnread: boolean
  warnings: string[]
}

/** Numero tolerante: "29,801.88" | "29.801,88" | "3,000.00" | "7.500" | "-" -> null/valor */
export function parseNum(raw: string | undefined | null): number | null {
  if (raw == null) return null
  let s = String(raw).trim().replace(/[€\s%]/g, "")
  if (!s || s === "-" || s === "–" || s.toLowerCase() === "null") return null
  const neg = /^\(.*\)$/.test(s) || s.startsWith("-")
  s = s.replace(/[()]/g, "").replace(/^-/, "")
  const lastComma = s.lastIndexOf(",")
  const lastDot = s.lastIndexOf(".")
  if (lastComma >= 0 && lastDot >= 0) {
    // el separador MAS A LA DERECHA es el decimal
    if (lastComma > lastDot) s = s.replace(/\./g, "").replace(",", ".")
    else s = s.replace(/,/g, "")
  } else if (lastComma >= 0) {
    // solo comas: decimal si va seguida de 1-2 digitos al final; si no, miles
    s = /,\d{1,2}$/.test(s) ? s.replace(/\./g, "").replace(",", ".") : s.replace(/,/g, "")
  }
  // solo puntos: "7.500" es ambiguo (7.5 o 7500). Regla: si UN punto con 3 decimales
  // exactos y el entero es corto (<=2 digitos), se trata como decimal (tipos 7.500%).
  // "29.801" (miles espanol) no llega aqui porque las cantidades traen 2 decimales.
  const n = Number.parseFloat(s)
  return Number.isNaN(n) ? null : neg ? -n : n
}

function parseDateToken(raw: string | undefined): string {
  if (!raw) return ""
  const s = raw.trim()
  let m = s.match(/^(\d{4})-(\d{1,2})-(\d{1,2})/)
  if (m) return `${m[1]}-${m[2].padStart(2, "0")}-${m[3].padStart(2, "0")}`
  m = s.match(/^(\d{1,2})[\/.-](\d{1,2})[\/.-](\d{4})/)
  if (m) return `${m[3]}-${m[2].padStart(2, "0")}-${m[1].padStart(2, "0")}`
  return ""
}

const DATE_RE = /\d{4}-\d{1,2}-\d{1,2}|\d{1,2}[\/.-]\d{1,2}[\/.-]\d{4}/
const NUM_RE = /-?[\d.,]+\d|\d/
const CHECK_RE = /[✓✔√]|(?:^|\s)(yes|si|sí|y)(?:\s|$)/i
const CROSS_RE = /[✗✘×xX]|(?:^|\s)(no|n)(?:\s|$)/i

function yn(token: string | undefined): "Y" | "N" | "" {
  if (!token) return ""
  if (CHECK_RE.test(token)) return "Y"
  if (CROSS_RE.test(token)) return "N"
  return ""
}

// Etiquetas conocidas de la pantalla: si la linea siguiente a una etiqueta VACIA es en
// realidad OTRA etiqueta, no hay que devolverla como si fuera el valor (evita robar el
// dato del campo contiguo cuando el propio campo viene vacio en 4Sight).
const STOP_LABELS =
  /Account Closure Date|Calculation Date|Interest Applicable Rate|Penalty Interest Applicable Rate|Judicial Costs|Updated Interest|Updated Penalty Interest|\bConsumer\b|First Residence|Total Updated Debt|^Capital\b|^Interest\b|Key Borrower|Fiscal Number|Loan Number|LegalID/i

/** Busca la etiqueta y devuelve el texto que la sigue (misma linea o hasta N lineas despues).
 *  Si la etiqueta esta vacia, NO cruza a una linea que sea a su vez otra etiqueta conocida
 *  (evita robar el valor del campo siguiente cuando el propio campo viene sin dato). */
function after(lines: string[], label: RegExp, lookahead = 1): string {
  for (let i = 0; i < lines.length; i++) {
    const m = lines[i].match(label)
    if (!m) continue
    const rest = lines[i].slice((m.index ?? 0) + m[0].length).trim()
    if (rest) return rest
    const candidates = lines.slice(i + 1, i + 1 + lookahead)
    if (candidates.some((l) => STOP_LABELS.test(l))) return ""
    return candidates.join(" ").trim()
  }
  return ""
}

function firstMatch(text: string, re: RegExp): string {
  const m = text.match(re)
  return m ? m[0] : ""
}

export function parse4Sight(text: string): Parsed4Sight {
  const warnings: string[] = []
  const clean = text.replace(/\r/g, "")
  const lines = clean.split("\n").map((l) => l.trim()).filter(Boolean)
  const flat = lines.join(" ")

  const out: Parsed4Sight = {
    loanNumber: "", legalId: "", folder: "", borrower: "",
    accountClosureDate: "", calculationDate: "",
    interestRatePct: null, penaltyRatePct: null,
    consumer: "", firstResidence: "",
    updatedInterest: null, judicialCosts: null,
    capital: null, interest: null, updatedPenaltyInterest: null,
    otherExpenses: null, totalUpdatedDebt: null, penaltyADE: null,
    properties: [], iconFieldsUnread: true, warnings,
  }

  // ---- expediente ----
  out.loanNumber = (flat.match(/Loan Number:?\s*([\w-]+)/i)?.[1] ?? "").trim()
  out.legalId = (flat.match(/LegalID=?\s*(\d+)/i)?.[1] ?? "").trim()
  out.folder = (flat.match(/\b([A-Z]{2,4}-[A-Z]{2}-\d{2}-\d{3,5})\b/)?.[1] ?? "").trim()
  const kb = after(lines, /Key Borrower/i, 1)
  out.borrower = kb.split(/Fiscal Number/i)[0].trim().slice(0, 80)

  // ---- Debt Calculation (etiquetas escalares) ----
  out.accountClosureDate = parseDateToken(firstMatch(after(lines, /Account Closure Date/i), DATE_RE))
  out.calculationDate = parseDateToken(firstMatch(after(lines, /Calculation Date/i), DATE_RE))
  out.interestRatePct = parseNum(firstMatch(after(lines, /(?<!Penalty )Interest Applicable Rate/i), NUM_RE))
  out.penaltyRatePct = parseNum(firstMatch(after(lines, /Penalty Interest Applicable Rate/i), NUM_RE))
  out.judicialCosts = parseNum(firstMatch(after(lines, /Judicial Costs(?! Maximum| Amount)/i), NUM_RE))
  out.updatedInterest = parseNum(firstMatch(after(lines, /Updated Interest(?! Amount)/i), NUM_RE))

  // Consumer / First Residence: simbolo tras la etiqueta
  out.consumer = yn(after(lines, /\bConsumer\b/i, 1).slice(0, 12))
  out.firstResidence = yn(after(lines, /First Residence/i, 1).slice(0, 12))

  // Caso tabla linealizada: fila de valores de Debt Calculation
  // "2022-12-15 2026-07-17 7.500 7.500 ✗ 8,022.01 ✗ 3,000.00"
  if (!out.accountClosureDate || out.interestRatePct == null) {
    for (const l of lines) {
      const dates = l.match(new RegExp(DATE_RE.source, "g"))
      if (dates && dates.length >= 2) {
        const nums = l.replace(new RegExp(DATE_RE.source, "g"), " ").match(/-?[\d.,]+\d/g) ?? []
        if (nums.length >= 3) {
          out.accountClosureDate = out.accountClosureDate || parseDateToken(dates[0])
          out.calculationDate = out.calculationDate || parseDateToken(dates[1])
          if (out.interestRatePct == null) out.interestRatePct = parseNum(nums[0])
          if (out.penaltyRatePct == null) out.penaltyRatePct = parseNum(nums[1])
          if (out.updatedInterest == null && nums.length >= 3) out.updatedInterest = parseNum(nums[2])
          if (out.judicialCosts == null && nums.length >= 4) out.judicialCosts = parseNum(nums[nums.length - 1])
          const symbols = l.match(/[✓✔√✗✘×]/g) ?? []
          if (!out.consumer && symbols[0]) out.consumer = yn(symbols[0])
          if (!out.firstResidence && symbols[1]) out.firstResidence = yn(symbols[1])
          break
        }
      }
    }
  }

  // ---- Total Updated Debt: Capital / Interest / Updated Penalty Interest / Judicial Costs / Other / Total ----
  const tudIdx = lines.findIndex((l) => /Total Updated Debt/i.test(l))
  if (tudIdx >= 0) {
    for (let i = tudIdx; i < Math.min(tudIdx + 12, lines.length); i++) {
      const nums = lines[i].match(/-?[\d.,]+\d|-/g)?.filter((t) => t !== "") ?? []
      const values = nums.map((t) => parseNum(t))
      const numeric = values.filter((v): v is number => v != null)
      if (numeric.length >= 4 && !/Maximum|per Property/i.test(lines[i])) {
        out.capital = values[0]
        out.interest = values[1]
        out.updatedPenaltyInterest = values[2]
        out.judicialCosts = values[3] ?? out.judicialCosts
        if (values.length >= 6) {
          out.otherExpenses = values[4]
          out.totalUpdatedDebt = values[5]
        } else if (values.length === 5) {
          out.totalUpdatedDebt = values[4]
        }
        break
      }
    }
  }
  if (out.capital == null) {
    out.capital = parseNum(firstMatch(after(lines, /^Capital\b/i), NUM_RE))
    out.interest = out.interest ?? parseNum(firstMatch(after(lines, /^Interest\b(?! Applicable| Maximum)/i), NUM_RE))
    out.updatedPenaltyInterest = out.updatedPenaltyInterest ?? parseNum(firstMatch(after(lines, /Updated Penalty Interest/i), NUM_RE))
  }

  // ---- Criterio estandar: demora a cierre = UPI - UI (el Excel recalcula la actualizacion) ----
  if (out.updatedPenaltyInterest != null && out.updatedInterest != null) {
    const ade = Math.round((out.updatedPenaltyInterest - out.updatedInterest) * 100) / 100
    out.penaltyADE = ade >= 0 ? ade : out.updatedPenaltyInterest
    if (ade < 0) warnings.push("Updated Penalty Interest < Updated Interest: revise la demora a cierre (D12).")
  } else if (out.updatedPenaltyInterest != null) {
    out.penaltyADE = out.updatedPenaltyInterest
    warnings.push("No se encontro 'Updated Interest': se usa el Updated Penalty Interest integro como demora a cierre (D12). Revise el criterio.")
  }

  // ---- Propiedades: filas "PropertyID num num num num ..." en seccion MSA ----
  const msaIdx = lines.findIndex((l) => /Maximum Secured Amount per Property/i.test(l))
  const searchLines = msaIdx >= 0 ? lines.slice(msaIdx) : lines
  for (const l of searchLines) {
    if (/^Total\b/i.test(l)) continue
    const m = l.match(/^(\d{4,9})\b(.*)$/)
    if (!m) continue
    const tokens = m[2].match(/-?[\d.,]+\d|-|–/g) ?? []
    if (tokens.length < 4) continue
    const vals = tokens.map((t) => parseNum(t))
    if (out.properties.some((p) => p.propertyId === m[1])) continue
    out.properties.push({
      propertyId: m[1],
      firstResidence: out.firstResidence,
      principalMSA: vals[0] ?? null,
      interestMSA: vals[1] ?? null,
      penaltyMSA: vals[2] ?? null,
      judicialMSA: vals[3] ?? null,
      otherMSA: vals.length >= 6 ? vals[4] : null, // col Other antes del Total cuando hay 6+
    })
    if (out.properties.length >= 30) {
      warnings.push("Mas de 30 fincas: la plantilla admite 30 (filas 30-59); se toman las 30 primeras.")
      break
    }
  }

  // ---- avisos de calidad ----
  // Consumer / First Residence en 4Sight son ICONOS (check/cruz): Ctrl+C no los copia
  // como texto, asi que se confirman a mano en el panel destacado de la app.
  out.iconFieldsUnread = out.consumer === "" || out.firstResidence === ""
  if (!out.capital) warnings.push("No se pudo leer el Capital: rellenelo a mano.")
  if (!out.accountClosureDate) warnings.push("No se pudo leer la Account Closure Date: rellenela a mano.")
  if (!out.calculationDate) warnings.push("No se pudo leer la Calculation Date: rellenela a mano.")
  if (out.properties.length === 0) warnings.push("No se detectaron fincas (PropertyID + responsabilidades MSA): anadalas a mano si las hay.")
  if (out.capital != null && out.interest != null && out.updatedPenaltyInterest != null && out.judicialCosts != null && out.totalUpdatedDebt != null) {
    const sum = Math.round((out.capital + out.interest + out.updatedPenaltyInterest + out.judicialCosts + (out.otherExpenses ?? 0)) * 100) / 100
    if (Math.abs(sum - out.totalUpdatedDebt) > 0.02) {
      warnings.push(`La suma de conceptos (${sum.toFixed(2)}) no cuadra con el Total Updated Debt (${out.totalUpdatedDebt.toFixed(2)}): revise los importes.`)
    }
  }
  return out
}
