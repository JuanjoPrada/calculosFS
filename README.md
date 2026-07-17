# calculosFS

Herramienta web de Legal Ops (Finsolutia): convierte los datos de la pantalla
**Debt Calculation / LegalUnpaidPrincipalBalance** de 4Sight en el fichero oficial
`Calculo_Deuda_y_Subasta.xlsx` con la pestaña **Informe Subasta** populada.

## Cómo se usa (3 pasos)

1. En 4Sight, con el préstamo abierto: `Ctrl + A` y `Ctrl + C`.
2. En calculosFS: botón **«Pegar datos de 4Sight»** y revisar los campos.
3. Botón **«Descargar Excel»**. El libro descargado conserva todas las fórmulas,
   formatos y hojas (Cálculo Intereses, Formato Cliente, CES, HO…) y se recalcula
   al abrirse.

Todo se procesa en el navegador: ningún dato sale del equipo.

## Criterios estándar aplicados

- `Principal → D10`, `Interest → D11`, `Judicial Costs → D14`, fechas → `D15/D16`,
  tipos (en %) → `D22/D23` (en fracción), `Consumer → D21`.
- **Demora a cierre (D12) = Updated Penalty Interest − Updated Interest**: el Excel
  recalcula la actualización desde la fecha de cierre, de modo que el total coincide
  con el Total Updated Debt de 4Sight.
- Fincas: `PropertyID → A`, `IsFirstResidence → B`, MSAs → `I/J/K` y
  `L = JudicialCosts MSA + OtherExpenses MSA`. El BaseAmount (tipo de subasta) no
  aparece en esa pantalla y puede completarse a mano.
- Defaults: `Mode = FIJO`, `NegativeRatesToZero = Sí`, `Ley 1/2025 = N (PRE)`,
  cargas 2ª/3ª = 0.

## Plantilla

La plantilla vive **congelada** en `public/plantilla/Calculo_Deuda_y_Subasta.xlsx`
(copia del libro individual del Proyecto Legal XXI). El populado edita únicamente
los valores de las celdas de entrada del Informe Subasta mediante manipulación
directa del XML (JSZip), por lo que las direcciones de celda son estables. Si el
libro maestro cambia de layout, hay que volver a congelar la plantilla y revisar
`lib/populate-xlsx.ts`.

## Despliegue (Netlify)

Site nuevo → Import from Git → este repositorio. Netlify detecta Next.js
(`netlify.toml` incluido: `next build`, Node 20). No requiere variables de entorno.
