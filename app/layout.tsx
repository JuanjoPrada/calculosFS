import type { Metadata } from "next"
import type { ReactNode } from "react"
import "./globals.css"

export const metadata: Metadata = {
  title: "calculosFS — Cálculo de deuda y subasta",
  description:
    "Pegue los datos de 4Sight y descargue el fichero de Cálculo de Deuda y Subasta con el Informe de Subasta populado.",
}

export default function RootLayout({ children }: { children: ReactNode }) {
  return (
    <html lang="es">
      <body className="min-h-screen bg-slate-100 text-slate-900 antialiased">{children}</body>
    </html>
  )
}
