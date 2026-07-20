"use client"

import { useMemo, useState } from "react"
import { parse4Sight, parseNum, type Parsed4Sight } from "@/lib/parse4sight"
import { populateAndDownload, type FincaInput, type InformeInput } from "@/lib/populate-xlsx"

type Status = { kind: "ok" | "error" | "info"; text: string } | null

const EMPTY: InformeInput = {
  loanNumber: "", legalId: "", comentarios: "",
  principal: "", interest: "", penaltyADE: "", expenses: "", judicialCosts: "",
  accountClosureDate: "", calculationDate: "",
  mode: "FIJO", consumer: "N", interestRatePct: "", penaltyRatePct: "",
  regimen12025: "N", executionOrderAmount: "", budgetAmount: "",
  fincas: [],
}

const nf = new Intl.NumberFormat("es-ES", { minimumFractionDigits: 2, maximumFractionDigits: 2 })

export default function Page() {
  const [form, setForm] = useState<InformeInput>(EMPTY)
  const [warnings, setWarnings] = useState<string[]>([])
  const [status, setStatus] = useState<Status>(null)
  const [hasData, setHasData] = useState(false)
  const [showManual, setShowManual] = useState(false)
  const [manualText, setManualText] = useState("")
  const [downloading, setDownloading] = useState(false)
  const [checkTotal, setCheckTotal] = useState<number | null>(null)
  const [needsConfirm, setNeedsConfirm] = useState(false)

  function applyParsed(p: Parsed4Sight) {
    setNeedsConfirm(p.iconFieldsUnread)
    const hoy = new Date().toISOString().slice(0, 10)
    setForm({
      loanNumber: p.loanNumber,
      legalId: p.legalId,
      comentarios: `Importado de 4Sight el ${hoy}${p.folder ? ` · ${p.folder}` : ""}${p.borrower ? ` · ${p.borrower}` : ""}`,
      principal: p.capital != null ? String(p.capital) : "",
      interest: p.interest != null ? String(p.interest) : "",
      penaltyADE: p.penaltyADE != null ? String(p.penaltyADE) : "",
      expenses: p.otherExpenses != null ? String(p.otherExpenses) : "0",
      judicialCosts: p.judicialCosts != null ? String(p.judicialCosts) : "0",
      accountClosureDate: p.accountClosureDate,
      calculationDate: p.calculationDate,
      mode: "FIJO",
      consumer: p.consumer || "N",
      interestRatePct: p.interestRatePct != null ? String(p.interestRatePct) : "",
      penaltyRatePct: p.penaltyRatePct != null ? String(p.penaltyRatePct) : "",
      regimen12025: "N",
      executionOrderAmount: "",
      budgetAmount: "",
      fincas: p.properties.map((pr) => ({
        propertyId: pr.propertyId,
        firstResidence: pr.firstResidence || p.firstResidence || "N",
        baseAmount: "",
        principalMSA: pr.principalMSA != null ? String(pr.principalMSA) : "",
        interestMSA: pr.interestMSA != null ? String(pr.interestMSA) : "",
        penaltyMSA: pr.penaltyMSA != null ? String(pr.penaltyMSA) : "",
        otherJudMSA: String(((pr.judicialMSA ?? 0) + (pr.otherMSA ?? 0)) || ""),
      })),
    })
    setWarnings(p.warnings)
    setCheckTotal(p.totalUpdatedDebt)
    setHasData(true)
    setStatus({ kind: "ok", text: "Datos importados. Revise los campos y pulse Descargar." })
  }

  async function handlePasteButton() {
    setStatus(null)
    try {
      if (!navigator.clipboard?.readText) throw new Error("clipboard")
      const text = await navigator.clipboard.readText()
      if (!text || text.trim().length < 40) {
        setStatus({ kind: "error", text: "El portapapeles está vacío. En 4Sight pulse Ctrl+A y después Ctrl+C, y vuelva a intentarlo." })
        setShowManual(true)
        return
      }
      applyParsed(parse4Sight(text))
    } catch {
      setShowManual(true)
      setStatus({ kind: "info", text: "Su navegador no permite leer el portapapeles directamente: pegue el texto en el recuadro de abajo (Ctrl+V) y pulse Procesar." })
    }
  }

  function handleManual() {
    if (manualText.trim().length < 40) {
      setStatus({ kind: "error", text: "Pegue primero el contenido copiado de 4Sight (Ctrl+V en el recuadro)." })
      return
    }
    applyParsed(parse4Sight(manualText))
  }

  async function handleDownload() {
    setDownloading(true)
    setStatus(null)
    try {
      await populateAndDownload(form)
      setStatus({ kind: "ok", text: "Excel descargado. Al abrirlo, pulse «Habilitar edición» si Excel lo pide: las fórmulas se recalculan solas." })
    } catch (e: any) {
      setStatus({ kind: "error", text: e?.message || "No se pudo generar el Excel." })
    } finally {
      setDownloading(false)
    }
  }

  function upd(field: keyof InformeInput, value: string) {
    setForm((f) => ({ ...f, [field]: value }))
  }
  function updFinca(i: number, field: keyof FincaInput, value: string) {
    setForm((f) => ({ ...f, fincas: f.fincas.map((x, j) => (j === i ? { ...x, [field]: value } : x)) }))
  }
  function addFinca() {
    setForm((f) => ({ ...f, fincas: [...f.fincas, { propertyId: "", firstResidence: "N", baseAmount: "", principalMSA: "", interestMSA: "", penaltyMSA: "", otherJudMSA: "" }] }))
  }
  function delFinca(i: number) {
    setForm((f) => ({ ...f, fincas: f.fincas.filter((_, j) => j !== i) }))
  }
  function setAllFirstResidence(v: string) {
    setForm((f) => ({ ...f, fincas: f.fincas.map((x) => ({ ...x, firstResidence: v })) }))
  }
  // Valor comun de "vivienda habitual" de las fincas ("" si estan mezcladas)
  const fincasFR = form.fincas.length === 0 ? "N" : form.fincas.every((x) => x.firstResidence === form.fincas[0].firstResidence) ? form.fincas[0].firstResidence : ""

  // Control de cuadre contra el Total Updated Debt de 4Sight
  const cuadre = useMemo(() => {
    if (checkTotal == null) return null
    const s = ["principal", "interest", "penaltyADE", "expenses", "judicialCosts"]
      .map((k) => parseNum(form[k as keyof InformeInput] as string) ?? 0)
      .reduce((a, b) => a + b, 0)
    // el Excel añade el interés actualizado (UI) que 4Sight ya incluye en el UPI
    return { esperado: checkTotal, conceptos: s }
  }, [form, checkTotal])

  return (
    <main className="mx-auto max-w-6xl px-4 py-6">
      <header className="mb-6 flex items-center gap-3">
        <div className="flex h-11 w-11 items-center justify-center rounded-xl bg-blue-700 text-xl font-black text-white shadow">FS</div>
        <div>
          <h1 className="text-2xl font-bold tracking-tight">calculosFS</h1>
          <p className="text-sm text-slate-500">De 4Sight al Excel de Cálculo de Deuda y Subasta en dos clics.</p>
        </div>
      </header>

      {/* PASO 1 */}
      <section className="rounded-2xl border border-slate-200 bg-white p-5 shadow-sm">
        <h2 className="mb-1 text-lg font-semibold"><span className="mr-2 inline-flex h-7 w-7 items-center justify-center rounded-full bg-blue-700 text-sm font-bold text-white">1</span>Copie la pantalla de 4Sight</h2>
        <ol className="mb-4 list-decimal space-y-1.5 pl-5 text-sm text-slate-600">
          <li>En 4Sight, calcule con el <strong>Simulador</strong> (pestaña Legal UPB → Simulator → Result).</li>
          <li>
            En el resultado, <strong>despliegue el panel «Loan Number: …»</strong> (pulse la flecha <span className="font-mono">›</span> para abrirlo):
            solo con el panel abierto se copian los datos de la fila (Account Closure Date, tipos, importes y fincas).
          </li>
          <li>
            Con el panel ya desplegado, pulse
            <kbd className="mx-1 rounded border border-slate-300 bg-slate-100 px-1.5 py-0.5 text-xs font-semibold">Ctrl + A</kbd>
            y después
            <kbd className="mx-1 rounded border border-slate-300 bg-slate-100 px-1.5 py-0.5 text-xs font-semibold">Ctrl + C</kbd>.
          </li>
          <li>Vuelva aquí y pulse el botón de abajo.</li>
        </ol>
        <button
          onClick={handlePasteButton}
          className="w-full rounded-xl bg-blue-700 px-6 py-4 text-lg font-bold text-white shadow-md transition hover:bg-blue-800"
        >
          📋 Pegar datos de 4Sight
        </button>
        <button onClick={() => setShowManual(!showManual)} className="mt-2 text-sm text-slate-500 underline hover:text-slate-700">
          ¿No funciona el botón? Pegue el texto a mano
        </button>
        {showManual && (
          <div className="mt-3 space-y-2">
            <textarea
              value={manualText}
              onChange={(e) => setManualText(e.target.value)}
              placeholder="Pegue aquí (Ctrl+V) todo el texto copiado de 4Sight…"
              className="h-32 w-full rounded-lg border border-slate-300 p-3 text-sm"
            />
            <button onClick={handleManual} className="rounded-lg bg-slate-800 px-4 py-2 text-sm font-semibold text-white hover:bg-slate-900">Procesar texto pegado</button>
          </div>
        )}
      </section>

      {status && (
        <p className={`mt-4 rounded-xl border px-4 py-3 text-sm font-medium ${status.kind === "ok" ? "border-emerald-200 bg-emerald-50 text-emerald-800" : status.kind === "error" ? "border-red-200 bg-red-50 text-red-700" : "border-blue-200 bg-blue-50 text-blue-800"}`}>
        {status.text}
        </p>
      )}

      {/* PASO 2 */}
      {hasData && (
        <section className="mt-4 rounded-2xl border border-slate-200 bg-white p-5 shadow-sm">
          <h2 className="mb-1 text-lg font-semibold"><span className="mr-2 inline-flex h-7 w-7 items-center justify-center rounded-full bg-blue-700 text-sm font-bold text-white">2</span>Revise los datos</h2>
          <p className="mb-4 text-sm text-slate-600">Todo es editable. Estos valores se escribirán en la pestaña <strong>Informe Subasta</strong>; el Excel recalcula el resto (interés actualizado, costas, pujas…).</p>

          {warnings.length > 0 && (
            <div className="mb-4 space-y-1 rounded-xl border border-amber-200 bg-amber-50 px-4 py-3 text-sm text-amber-800">
              {warnings.map((w, i) => (<p key={i}>⚠ {w}</p>))}
            </div>
          )}

          {needsConfirm && (
            <div className="mb-4 rounded-xl border-2 border-amber-400 bg-amber-50 px-4 py-4">
              <p className="mb-1 text-sm font-bold text-amber-900">✋ Confirme estos 2 campos antes de descargar</p>
              <p className="mb-3 text-xs text-amber-800">
                En 4Sight <strong>«Consumer»</strong> y <strong>«First Residence»</strong> aparecen como iconos
                <span className="mx-1 font-semibold text-emerald-600">✓</span>/<span className="mx-1 font-semibold text-red-600">✗</span>
                y <strong>no se copian con Ctrl+C</strong>. Márquelos aquí según lo que vea en 4Sight:
              </p>
              <div className="flex flex-wrap gap-8">
                <div>
                  <span className="mb-1 block text-xs font-semibold uppercase tracking-wide text-amber-800">Consumidor</span>
                  <Toggle value={form.consumer} onChange={(v) => { upd("consumer", v); }} />
                </div>
                <div>
                  <span className="mb-1 block text-xs font-semibold uppercase tracking-wide text-amber-800">Vivienda habitual {form.fincas.length > 1 ? "(todas las fincas)" : ""}</span>
                  <Toggle value={fincasFR} onChange={setAllFirstResidence} />
                  {form.fincas.length > 1 && <span className="mt-1 block text-[11px] text-amber-700">Puede afinarla finca a finca en la tabla de abajo.</span>}
                </div>
              </div>
              <button onClick={() => setNeedsConfirm(false)} className="mt-3 text-xs font-medium text-amber-700 underline hover:text-amber-900">Confirmado, ocultar aviso</button>
            </div>
          )}

          <div className="grid grid-cols-1 gap-x-6 gap-y-3 sm:grid-cols-2 lg:grid-cols-4">
            <Field label="Loan Number" v={form.loanNumber} on={(v) => upd("loanNumber", v)} />
            <Field label="LegalID" v={form.legalId} on={(v) => upd("legalId", v)} hint="Si copió también la barra de direcciones, se rellena solo." />
            <Field label="Fecha cierre (Account Closure)" v={form.accountClosureDate} on={(v) => upd("accountClosureDate", v)} hint="AAAA-MM-DD" />
            <Field label="Fecha de cálculo" v={form.calculationDate} on={(v) => upd("calculationDate", v)} hint="AAAA-MM-DD" />
            <Field label="Principal (€) → D10" v={form.principal} on={(v) => upd("principal", v)} />
            <Field label="Interest (€) → D11" v={form.interest} on={(v) => upd("interest", v)} />
            <Field label="Demora a cierre (€) → D12" v={form.penaltyADE} on={(v) => upd("penaltyADE", v)} hint="Criterio estándar: Updated Penalty − Updated Interest (el Excel recalcula la actualización)." />
            <Field label="Expenses (€) → D13" v={form.expenses} on={(v) => upd("expenses", v)} />
            <Field label="Judicial Costs (€) → D14" v={form.judicialCosts} on={(v) => upd("judicialCosts", v)} />
            <Field label="Tipo interés (%) → D22" v={form.interestRatePct} on={(v) => upd("interestRatePct", v)} />
            <Field label="Tipo demora (%) → D23" v={form.penaltyRatePct} on={(v) => upd("penaltyRatePct", v)} />
            <Select label="Modo de interés" v={form.mode} on={(v) => upd("mode", v)} opts={["FIJO", "VARIABLE", "LEGAL"]} />
            <Select label="Consumidor" v={form.consumer} on={(v) => upd("consumer", v)} opts={["N", "Y"]} />
            <Select label="Ley 1/2025 (régimen subasta)" v={form.regimen12025} on={(v) => upd("regimen12025", v)} opts={["N", "Y"]} hint="N = procedimiento anterior (PRE), Y = posterior (POST)." />
            <Field label="ExecutionOrderAmount (€)" v={form.executionOrderAmount} on={(v) => upd("executionOrderAmount", v)} hint="Opcional; no viene en esta pantalla." />
            <Field label="BudgetAmount (€)" v={form.budgetAmount} on={(v) => upd("budgetAmount", v)} hint="Opcional." />
          </div>

          {cuadre && (
            <p className="mt-4 rounded-lg bg-slate-50 px-4 py-2 text-sm text-slate-600">
              Control: 4Sight declara un total de <strong>{nf.format(cuadre.esperado)} €</strong>. El Excel llegará a esa cifra sumando a los conceptos el interés actualizado que recalcula solo.
            </p>
          )}

          <h3 className="mb-2 mt-6 text-base font-semibold">Fincas (responsabilidad hipotecaria máxima)</h3>
          <div className="overflow-x-auto rounded-xl border border-slate-200">
            <table className="w-full border-collapse text-sm">
              <thead>
                <tr className="bg-slate-50 text-left text-xs uppercase tracking-wide text-slate-500">
                  <th className="px-2 py-2">PropertyID</th>
                  <th className="px-2 py-2">Viv. habitual</th>
                  <th className="px-2 py-2">Tipo subasta (€)</th>
                  <th className="px-2 py-2">Principal MSA</th>
                  <th className="px-2 py-2">Interest MSA</th>
                  <th className="px-2 py-2">Penalty MSA</th>
                  <th className="px-2 py-2">Otros+Costas MSA</th>
                  <th className="px-2 py-2"></th>
                </tr>
              </thead>
              <tbody className="divide-y divide-slate-100">
                {form.fincas.map((f, i) => (
                  <tr key={i}>
                    <td className="px-2 py-1"><input className={cellCls} value={f.propertyId} onChange={(e) => updFinca(i, "propertyId", e.target.value)} /></td>
                    <td className="px-2 py-1">
                      <select className={cellCls} value={f.firstResidence} onChange={(e) => updFinca(i, "firstResidence", e.target.value)}>
                        <option value="N">No</option><option value="Y">Sí</option>
                      </select>
                    </td>
                    <td className="px-2 py-1"><input className={cellCls} value={f.baseAmount} onChange={(e) => updFinca(i, "baseAmount", e.target.value)} placeholder="(rellenar)" /></td>
                    <td className="px-2 py-1"><input className={cellCls} value={f.principalMSA} onChange={(e) => updFinca(i, "principalMSA", e.target.value)} /></td>
                    <td className="px-2 py-1"><input className={cellCls} value={f.interestMSA} onChange={(e) => updFinca(i, "interestMSA", e.target.value)} /></td>
                    <td className="px-2 py-1"><input className={cellCls} value={f.penaltyMSA} onChange={(e) => updFinca(i, "penaltyMSA", e.target.value)} /></td>
                    <td className="px-2 py-1"><input className={cellCls} value={f.otherJudMSA} onChange={(e) => updFinca(i, "otherJudMSA", e.target.value)} /></td>
                    <td className="px-1 py-1"><button onClick={() => delFinca(i)} className="text-slate-400 hover:text-red-600" aria-label="Eliminar finca">✕</button></td>
                  </tr>
                ))}
              </tbody>
            </table>
          </div>
          <button onClick={addFinca} className="mt-2 rounded-lg border border-slate-300 px-3 py-1.5 text-sm font-medium text-slate-600 hover:bg-slate-50">+ Añadir finca</button>
          <p className="mt-2 text-xs text-slate-400">El «Tipo de subasta» (BaseAmount) no aparece en esta pantalla de 4Sight: rellénelo si lo conoce; puede dejarse vacío y completarlo en el Excel.</p>

          <Field label="Comentarios (→ celda de comentarios del informe)" v={form.comentarios} on={(v) => upd("comentarios", v)} wide className="mt-4" />
        </section>
      )}

      {/* PASO 3 */}
      {hasData && (
        <section className="mt-4 rounded-2xl border border-slate-200 bg-white p-5 shadow-sm">
          <h2 className="mb-3 text-lg font-semibold"><span className="mr-2 inline-flex h-7 w-7 items-center justify-center rounded-full bg-blue-700 text-sm font-bold text-white">3</span>Descargue el Excel</h2>
          <button
            onClick={handleDownload}
            disabled={downloading}
            className="w-full rounded-xl bg-emerald-600 px-6 py-4 text-lg font-bold text-white shadow-md transition hover:bg-emerald-700 disabled:cursor-not-allowed disabled:opacity-60"
          >
            {downloading ? "Generando…" : "⬇ Descargar Calculo_Deuda_y_Subasta.xlsx"}
          </button>
          <p className="mt-3 text-xs text-slate-400">
            El fichero es la plantilla oficial con la pestaña «Informe Subasta» populada. Todas las fórmulas, formatos y hojas
            (Cálculo Intereses, Formato Cliente, CES, HO…) permanecen intactas y se recalculan al abrir el libro.
          </p>
        </section>
      )}

      <footer className="mt-6 text-center text-xs text-slate-400">
        <p>calculosFS · Legal Ops Finsolutia · Los datos se procesan íntegramente en su navegador (no se envían a ningún servidor).</p>
      </footer>
    </main>
  )
}

const cellCls = "w-full min-w-[7rem] rounded border border-slate-300 px-2 py-1 text-sm"

function Field({ label, v, on, hint, wide, className = "" }: { label: string; v: string; on: (v: string) => void; hint?: string; wide?: boolean; className?: string }) {
  return (
    <label className={`block ${wide ? "sm:col-span-2 lg:col-span-4" : ""} ${className}`}>
      <span className="mb-0.5 block text-xs font-semibold uppercase tracking-wide text-slate-500">{label}</span>
      <input className="w-full rounded-lg border border-slate-300 bg-white px-3 py-2 text-sm shadow-sm outline-none focus:border-blue-500 focus:ring-2 focus:ring-blue-100" value={v} onChange={(e) => on(e.target.value)} />
      {hint && <span className="mt-0.5 block text-[11px] leading-tight text-slate-400">{hint}</span>}
    </label>
  )
}

// Selector segmentado Sí/No (valor "" = sin definir / mixto).
function Toggle({ value, onChange }: { value: string; onChange: (v: string) => void }) {
  const btn = (v: string, label: string, active: boolean, color: string) =>
    <button type="button" onClick={() => onChange(v)}
      className={`px-5 py-2 text-sm font-bold transition ${active ? `${color} text-white` : "bg-white text-slate-600 hover:bg-slate-50"}`}>{label}</button>
  return (
    <div className="inline-flex overflow-hidden rounded-lg border border-slate-300">
      {btn("N", "No", value === "N", "bg-slate-700")}
      {btn("Y", "Sí", value === "Y", "bg-emerald-600")}
    </div>
  )
}

function Select({ label, v, on, opts, hint }: { label: string; v: string; on: (v: string) => void; opts: string[]; hint?: string }) {
  return (
    <label className="block">
      <span className="mb-0.5 block text-xs font-semibold uppercase tracking-wide text-slate-500">{label}</span>
      <select className="w-full rounded-lg border border-slate-300 bg-white px-3 py-2 text-sm shadow-sm outline-none focus:border-blue-500" value={v} onChange={(e) => on(e.target.value)}>
        {opts.map((o) => (<option key={o} value={o}>{o}</option>))}
      </select>
      {hint && <span className="mt-0.5 block text-[11px] leading-tight text-slate-400">{hint}</span>}
    </label>
  )
}
