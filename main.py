import requests
import pandas as pd
import tkinter as tk
from tkinter import filedialog, simpledialog, messagebox, ttk
import os
import threading

def mostrar_progreso(texto="Procesando..."):
    ventana = tk.Toplevel()
    ventana.title("Cargando...")
    ventana.geometry("400x120")
    ventana.resizable(False, False)

    label = tk.Label(ventana, text=texto, font=("Arial", 11))
    label.pack(pady=10)

    barra = ttk.Progressbar(ventana, mode='indeterminate')
    barra.pack(fill='x', padx=20, pady=10)
    barra.start(10)

    ventana.update()
    return ventana, barra


def main():
    root = tk.Tk()
    root.withdraw()

    # === 1️⃣ Pedir usuario y contraseña ===
    email = simpledialog.askstring("Inicio de sesión", "Correo electrónico:", parent=root)
    password = simpledialog.askstring("Inicio de sesión", "Contraseña:", show="*", parent=root)

    if not email or not password:
        messagebox.showerror("Error", "Debes ingresar tu correo y contraseña.")
        return

    progreso, barra = mostrar_progreso("Autenticando usuario...")

    # === 2️⃣ Obtener Bearer Token ===
    login_url = "https://universal-console-server-b7agk5thba-uc.a.run.app/login"
    login_body = {"email": email, "password": password}

    try:
        login_response = requests.post(login_url, json=login_body,verify=False)
        login_response.raise_for_status()
        token = login_response.json().get("token")
        if not token:
            raise Exception("No se recibió token")
    except Exception as e:
        progreso.destroy()
        messagebox.showerror("Error", f"No se pudo autenticar:\n{e}")
        return

    headers = {"Authorization": f"Bearer {token}", "Content-Type": "application/json"}

    progreso.destroy()

    # === 3️⃣ Seleccionar archivo CRM ===
    crm_path = filedialog.askopenfilename(
        title="Selecciona el archivo CRM.xlsx",
        filetypes=[("Archivos Excel", "*.xlsx *.xls")]
    )

    if not crm_path:
        messagebox.showerror("Error", "No seleccionaste un archivo CRM.")
        return

    progreso, barra = mostrar_progreso("Leyendo archivo CRM...")

    # === 4️⃣ Leer CRM ===
    try:
        crm_df = pd.read_excel(crm_path)
        crm_df.columns = crm_df.columns.str.strip()
        if "Serie" not in crm_df.columns:
            raise Exception("El archivo no contiene columna 'Serie'")
        series_crm = crm_df["Serie"].astype(str).unique().tolist()
    except Exception as e:
        progreso.destroy()
        messagebox.showerror("Error", f"No se pudo leer CRM:\n{e}")
        return

    progreso.destroy()

    # === 5️⃣ Llamar a /vaultlist para filtrar series ===
    progreso, barra = mostrar_progreso("Filtrando series activas en vaultlist...")

    try:
        vault_url = "https://universal-console-server-b7agk5thba-uc.a.run.app/vaultlist"
        vault_body = {
            "customer": "KOF",
            "path": [],
            "page_number": 1,
            "page_size": len(series_crm),
            "filter_by": series_crm
        }

        vault_response = requests.post(vault_url, headers=headers, json=vault_body,verify=False)
        vault_response.raise_for_status()
        vault_data = vault_response.json()

        # Excluir series con estatus != False
        series_excluir = [
            item["serial_number"]
            for item in vault_data
            if item.get("estatus") not in [False, None]
        ]
    except Exception as e:
        progreso.destroy()
        messagebox.showerror("Error", f"No se pudo consultar vaultlist:\n{e}")
        return

    progreso.destroy()

    # === 6️⃣ Obtener datos de alerts_drawer ===
    progreso, barra = mostrar_progreso("Descargando alertas y fallas desde API...")

    api_url = "https://universal-console-server-b7agk5thba-uc.a.run.app/alerts_drawer"
    body = {
        "customer": "KOF",
        "class": "OPE",
        "algorithm": [
            "COMPRESSOR_FAIL",
            "TEMPERATURE_FAIL",
            "COMPRESSOR_RUN_TIME_EXCEEDED_ALERT"
        ],
        "path": [],
        "page_size": 500000,
        "page_number": 1
    }

    try:
        response = requests.post(api_url, headers=headers, json=body,verify=False)
        response.raise_for_status()
        data = response.json()
    except Exception as e:
        progreso.destroy()
        messagebox.showerror("Error", f"No se pudo obtener datos de alertas:\n{e}")
        return

    progreso.destroy()

    # === 7️⃣ Procesar resultados ===
    progreso, barra = mostrar_progreso("Procesando y filtrando resultados...")

    def pasa_filtros(item):
        if item.get("ControlDeActivos", "").strip().lower() != "sin riesgo":
            return False
        if item.get("EstatusKOF", "").strip().upper() != "LEGL":
            return False
        for campo in ["CodigoPostal", "EntreCalles", "DirecciónPdV", "PdV", "IdPdV"]:
            if item.get(campo, "").strip() != "":
                return False
        return True

    fallas, alertas = [], []
    for item in data:
        serie = str(item.get("Serie", ""))
        falla_alerta = item.get("Estatus", "").lower()

        if (
            serie not in series_crm and
            serie not in series_excluir and
            pasa_filtros(item)
        ):
            if "falla" in falla_alerta:
                fallas.append(item)
            elif "alerta" in falla_alerta or "demanda" in falla_alerta:
                alertas.append(item)

    progreso.destroy()

    # === 8️⃣ Guardar archivos ===
    save_dir = filedialog.askdirectory(title="Selecciona carpeta para guardar resultados")
    if not save_dir:
        messagebox.showinfo("Cancelado", "No seleccionaste carpeta.")
        return

    saved_files = []
    if fallas:
        fallas_path = os.path.join(save_dir, "Fallas_Nuevas.xlsx")
        pd.DataFrame(fallas).to_excel(fallas_path, index=False)
        saved_files.append(fallas_path)

    if alertas:
        alertas_path = os.path.join(save_dir, "Alertas_Nuevas.xlsx")
        pd.DataFrame(alertas).to_excel(alertas_path, index=False)
        saved_files.append(alertas_path)

    if saved_files:
        messagebox.showinfo("Completado", "✅ Archivos generados:\n\n" + "\n".join(saved_files))
    else:
        messagebox.showinfo("Sin nuevos registros", "No se encontraron alertas o fallas fuera del CRM.")


if __name__ == "__main__":
    main()
