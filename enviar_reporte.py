"""
=========================================================
  AUTOMATIZACIÓN DE REPORTES POR EMAIL — Javier Onel
  Versión 1.0 | github.com/Srpython1983
=========================================================
Lee cualquier archivo Excel o CSV y envía un reporte
por email automáticamente.

Instalación:
  pip install pandas openpyxl rich

Uso:
  python enviar_reporte.py
  python enviar_reporte.py mi_archivo.xlsx
"""

import os
import sys
import smtplib
import ssl
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from email.mime.base import MIMEBase
from email import encoders
from datetime import datetime

# ── Verificar dependencias ──────────────────────────────
try:
    import pandas as pd
    from rich.console import Console
    from rich.panel import Panel
    from rich.table import Table
    from rich import box
except ImportError:
    print("\n⚠️  Instala las dependencias:")
    print("   pip install pandas openpyxl rich\n")
    sys.exit(1)

console = Console()

# ═══════════════════════════════════════════════════════
#  CONFIGURACIÓN — EDITA AQUÍ TUS DATOS
# ═══════════════════════════════════════════════════════

CONFIG = {
    # ── Tu email de envío ──────────────────────────────
    "email_remitente":  "tu_email@gmail.com",
    "nombre_remitente": "Javier Onel",

    # ── Contraseña de aplicación ───────────────────────
    # Gmail: ve a Seguridad → Verificación en 2 pasos →
    # Contraseñas de aplicación → Generar
    "password": "xxxx xxxx xxxx xxxx",

    # ── Servidor SMTP ──────────────────────────────────
    # Gmail:   smtp.gmail.com  puerto 465
    # Outlook: smtp.office365.com  puerto 587
    # Yahoo:   smtp.mail.yahoo.com  puerto 465
    "smtp_servidor": "smtp.gmail.com",
    "smtp_puerto":   465,

    # ── Destinatarios ─────────────────────────────────
    # Agrega todos los emails que quieras
    "destinatarios": [
        "destinatario1@empresa.com",
        "destinatario2@empresa.com",
    ],

    # ── Asunto del email ───────────────────────────────
    "asunto": "Reporte Automático — {fecha}",

    # ── Empresa ───────────────────────────────────────
    "empresa": "Mi Empresa SpA",
}

# ═══════════════════════════════════════════════════════
#  LECTORES DE ARCHIVOS
# ═══════════════════════════════════════════════════════

def leer_excel(ruta):
    df_dict = pd.read_excel(ruta, sheet_name=None)
    datos = {}
    for hoja, df in df_dict.items():
        df = df.dropna(how="all").fillna("")
        datos[hoja] = df
    return datos

def leer_csv(ruta):
    for enc in ["utf-8-sig", "latin-1", "utf-16"]:
        try:
            df = pd.read_csv(ruta, encoding=enc)
            return {"Datos": df.fillna("")}
        except Exception:
            continue
    raise ValueError("No se pudo leer el CSV.")

def leer_archivo(ruta):
    ext = os.path.splitext(ruta)[1].lower()
    if ext in [".xlsx", ".xls"]:
        return leer_excel(ruta)
    elif ext == ".csv":
        return leer_csv(ruta)
    else:
        raise ValueError(f"Formato '{ext}' no soportado. Usa Excel o CSV.")

# ═══════════════════════════════════════════════════════
#  GENERAR REPORTE EN TEXTO
# ═══════════════════════════════════════════════════════

def generar_reporte_texto(datos, nombre_archivo):
    fecha = datetime.now().strftime("%d/%m/%Y %H:%M")
    lineas = []
    lineas.append("=" * 60)
    lineas.append(f"  REPORTE AUTOMÁTICO — {CONFIG['empresa']}")
    lineas.append(f"  Archivo: {nombre_archivo}")
    lineas.append(f"  Generado: {fecha}")
    lineas.append("=" * 60)
    lineas.append("")

    for hoja, df in datos.items():
        lineas.append(f"[ {hoja} ]")
        lineas.append("-" * 40)

        # Resumen de números
        nums = df.select_dtypes(include="number")
        if not nums.empty:
            lineas.append("Resumen numérico:")
            for col in nums.columns:
                total = nums[col].sum()
                promedio = nums[col].mean()
                lineas.append(f"  {col}:")
                lineas.append(f"    Total:   {total:,.2f}")
                lineas.append(f"    Promedio: {promedio:,.2f}")
                lineas.append(f"    Registros: {nums[col].count()}")
            lineas.append("")

        # Primeras 10 filas
        lineas.append(f"Primeras filas ({min(10, len(df))} de {len(df)} registros):")
        lineas.append(df.head(10).to_string(index=False))
        lineas.append("")

    lineas.append("=" * 60)
    lineas.append("Reporte generado automáticamente por:")
    lineas.append(f"github.com/Srpython1983 | WhatsApp: +56 9 9179 1778")
    lineas.append("=" * 60)

    return "\n".join(lineas)

def generar_reporte_html(datos, nombre_archivo):
    fecha = datetime.now().strftime("%d/%m/%Y %H:%M")

    html = f"""
    <html>
    <body style="font-family: Arial, sans-serif; color: #333; max-width: 800px; margin: 0 auto;">
        <div style="background: #1F3A8A; color: white; padding: 20px; border-radius: 8px 8px 0 0;">
            <h2 style="margin:0;">📊 Reporte Automático</h2>
            <p style="margin:5px 0 0;">{CONFIG['empresa']} — {fecha}</p>
        </div>
        <div style="background: #f5f5f5; padding: 15px; border-radius: 0 0 8px 8px;">
            <p><strong>Archivo analizado:</strong> {nombre_archivo}</p>
    """

    for hoja, df in datos.items():
        html += f"<h3 style='color:#1F3A8A;'>📋 {hoja}</h3>"

        nums = df.select_dtypes(include="number")
        if not nums.empty:
            html += "<table style='width:100%; border-collapse:collapse; margin-bottom:15px;'>"
            html += "<tr style='background:#1F3A8A; color:white;'>"
            html += "<th style='padding:8px; text-align:left;'>Columna</th>"
            html += "<th style='padding:8px; text-align:right;'>Total</th>"
            html += "<th style='padding:8px; text-align:right;'>Promedio</th>"
            html += "<th style='padding:8px; text-align:right;'>Registros</th>"
            html += "</tr>"

            for i, col in enumerate(nums.columns):
                bg = "#f9f9f9" if i % 2 == 0 else "#ffffff"
                total = nums[col].sum()
                prom = nums[col].mean()
                count = nums[col].count()
                html += f"<tr style='background:{bg};'>"
                html += f"<td style='padding:8px;'>{col}</td>"
                html += f"<td style='padding:8px; text-align:right;'>{total:,.2f}</td>"
                html += f"<td style='padding:8px; text-align:right;'>{prom:,.2f}</td>"
                html += f"<td style='padding:8px; text-align:right;'>{count}</td>"
                html += "</tr>"
            html += "</table>"

        html += f"<p style='color:#666; font-size:12px;'>Total de registros: {len(df)}</p>"

    html += """
        </div>
        <div style="text-align:center; padding:15px; color:#888; font-size:12px;">
            Reporte generado automáticamente por Javier Onel<br>
            <a href="https://github.com/Srpython1983">github.com/Srpython1983</a>
        </div>
    </body>
    </html>
    """
    return html

# ═══════════════════════════════════════════════════════
#  ENVIAR EMAIL
# ═══════════════════════════════════════════════════════

def enviar_email(reporte_texto, reporte_html, ruta_archivo, nombre_archivo):
    fecha = datetime.now().strftime("%d/%m/%Y")
    asunto = CONFIG["asunto"].format(
        fecha=fecha,
        empresa=CONFIG["empresa"]
    )

    msg = MIMEMultipart("alternative")
    msg["Subject"] = asunto
    msg["From"]    = f"{CONFIG['nombre_remitente']} <{CONFIG['email_remitente']}>"
    msg["To"]      = ", ".join(CONFIG["destinatarios"])

    msg.attach(MIMEText(reporte_texto, "plain", "utf-8"))
    msg.attach(MIMEText(reporte_html,  "html",  "utf-8"))

    # Adjuntar archivo original
    with open(ruta_archivo, "rb") as f:
        adjunto = MIMEBase("application", "octet-stream")
        adjunto.set_payload(f.read())
        encoders.encode_base64(adjunto)
        adjunto.add_header(
            "Content-Disposition",
            f"attachment; filename={nombre_archivo}"
        )
        msg.attach(adjunto)

    # Enviar
    context = ssl.create_default_context()
    with smtplib.SMTP_SSL(CONFIG["smtp_servidor"], CONFIG["smtp_puerto"], context=context) as server:
        server.login(CONFIG["email_remitente"], CONFIG["password"])
        server.sendmail(
            CONFIG["email_remitente"],
            CONFIG["destinatarios"],
            msg.as_string()
        )

# ═══════════════════════════════════════════════════════
#  MOSTRAR PREVIEW EN TERMINAL
# ═══════════════════════════════════════════════════════

def mostrar_preview(datos, nombre_archivo):
    console.print()
    console.print(Panel.fit(
        f"[bold]Archivo:[/bold] {nombre_archivo}\n"
        f"[bold]Hojas:[/bold] {', '.join(datos.keys())}",
        title="Vista previa del reporte",
        border_style="blue"
    ))

    for hoja, df in datos.items():
        nums = df.select_dtypes(include="number")
        if not nums.empty:
            tabla = Table(title=f"Resumen — {hoja}", box=box.ROUNDED, border_style="blue")
            tabla.add_column("Columna", style="bold")
            tabla.add_column("Total", justify="right")
            tabla.add_column("Promedio", justify="right")
            tabla.add_column("Registros", justify="right")

            for col in nums.columns:
                tabla.add_row(
                    col,
                    f"{nums[col].sum():,.2f}",
                    f"{nums[col].mean():,.2f}",
                    str(nums[col].count())
                )
            console.print()
            console.print(tabla)

# ═══════════════════════════════════════════════════════
#  MAIN
# ═══════════════════════════════════════════════════════

def main():
    console.print()
    console.print(Panel.fit(
        "[bold]AUTOMATIZACIÓN DE REPORTES POR EMAIL[/bold]\n"
        "[dim]by Javier Onel — github.com/Srpython1983[/dim]",
        border_style="blue"
    ))
    console.print()

    # Verificar configuración
    if CONFIG["email_remitente"] == "tu_email@gmail.com":
        console.print("[bold yellow]⚠️  Configura tu email en el archivo config[/bold yellow]")
        console.print("[dim]Edita la sección CONFIG al inicio del script[/dim]\n")

    # Obtener archivo
    if len(sys.argv) > 1:
        ruta = sys.argv[1]
    else:
        ruta = input("📂 Ruta del archivo (Excel o CSV): ").strip().strip("'\"")

    if not os.path.exists(ruta):
        console.print(f"[bold red]❌ Archivo no encontrado: {ruta}[/bold red]")
        sys.exit(1)

    nombre = os.path.basename(ruta)
    console.print(f"[dim]Leyendo: {nombre}[/dim]")

    try:
        datos = leer_archivo(ruta)
        console.print(f"[bold green]✅ Archivo leído — {len(datos)} hoja(s)[/bold green]")
    except Exception as e:
        console.print(f"[bold red]❌ Error: {e}[/bold red]")
        sys.exit(1)

    # Mostrar preview
    mostrar_preview(datos, nombre)

    # Confirmar envío
    console.print()
    console.print(f"[bold]Destinatarios:[/bold] {', '.join(CONFIG['destinatarios'])}")
    console.print(f"[bold]Desde:[/bold] {CONFIG['email_remitente']}")
    console.print()

    enviar = input("📧 ¿Enviar reporte por email? (s/n): ").strip().lower()

    if enviar == "s":
        try:
            console.print("\n[bold blue]Generando reporte...[/bold blue]")
            reporte_txt  = generar_reporte_texto(datos, nombre)
            reporte_html = generar_reporte_html(datos, nombre)

            console.print("[bold blue]Enviando email...[/bold blue]")
            enviar_email(reporte_txt, reporte_html, ruta, nombre)

            console.print(f"\n[bold green]✅ Reporte enviado exitosamente a:[/bold green]")
            for dest in CONFIG["destinatarios"]:
                console.print(f"   [green]• {dest}[/green]")
        except Exception as e:
            console.print(f"\n[bold red]❌ Error al enviar: {e}[/bold red]")
            console.print("[dim]Verifica tu email y contraseña de aplicación en CONFIG[/dim]")
    else:
        # Guardar reporte localmente
        guardar = input("💾 ¿Guardar reporte en archivo? (s/n): ").strip().lower()
        if guardar == "s":
            fecha   = datetime.now().strftime("%Y%m%d_%H%M%S")
            nombre_rep = f"reporte_{fecha}.txt"
            reporte_txt = generar_reporte_texto(datos, nombre)
            with open(nombre_rep, "w", encoding="utf-8") as f:
                f.write(reporte_txt)
            console.print(f"\n[bold green]✅ Reporte guardado: {nombre_rep}[/bold green]")

    console.print()
    console.print(Panel.fit(
        "[bold green]Proceso completado[/bold green]\n"
        "[dim]github.com/Srpython1983[/dim]",
        border_style="green"
    ))

if __name__ == "__main__":
    main()
