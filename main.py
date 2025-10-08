import requests
import pandas as pd
import tkinter as tk
from tkinter import filedialog, simpledialog, messagebox
import os

def main():
    root = tk.Tk()
    root.withdraw()  # Oculta la ventana principal

    # === 1️⃣ Pedir usuario y contraseña ===
    email = simpledialog.askstring("Inicio de sesión", "Correo electrónico:", parent=root)
    password = simpledialog.askstring("Inicio de sesión", "Contraseña:", show="*", parent=root)

    if not email or not password:
        messagebox.showerror("Error", "Debes ingresar tu correo y contraseña.")
        return

    # === 2️⃣ Obtener Bearer Token desde la API de login ===
    login_url = "https://universal-console-server-b7agk5thba-uc.a.run.app/login"
    login_body = {"email": email, "password": password}

    try:
        login_response = requests.post(login_url, json=login_body)
        if login_response.status_code != 200:
            messagebox.showerror("Error", f"Error de autenticación ({login_response.status_code})")
            return
        token_data = login_response.json()
        token = token_data.get("token")
        if not token:
            messagebox.showerror("Error", "No se recibió el token en la respuesta.")
            return
    except Exception as e:
        messagebox.showerror("Error", f"No se pudo conectar a la API de login:\n{e}")
        return

    # === 3️⃣ Seleccionar archivo CRM ===
    crm_path = filedialog.askopenfilename(
        title="Selecciona el archivo CRM.xlsx",
        filetypes=[("Archivos Excel", "*.xlsx *.xls")]
    )

    if not crm_path:
        messagebox.showerror("Error", "No seleccionaste un archivo CRM.")
        return

    # === 4️⃣ Hacer petición a la API alerts_drawer ===
    api_url = "https://universal-console-server-b7agk5thba-uc.a.run.app/alerts_drawer"
    headers = {"Authorization": f"Bearer {token}", "Content-Type": "application/json"}
    body = {
        "customer": "KOF",
        "class": "OPE",
        "algorithm": ["COMPRESSOR_FAIL","TEMPERATURE_FAIL","COMPRESSOR_RUN_TIME_EXCEEDED_ALERT"],
        "path": [],
        "page_size": 414491,
        "page_number": 1
    }

    try:
        response = requests.post(api_url, headers=headers, json=body)
        response.raise_for_status()
        data = response.json()
    except Exception as e:
        messagebox.showerror("Error", f"No se pudo obtener datos del API alerts_drawer:\n{e}")
        return

    # === 5️⃣ Leer CRM y comparar Series ===
    try:
        crm_df = pd.read_excel(crm_path)
        crm_df.columns = crm_df.columns.str.strip()
        if "Serie" not in crm_df.columns:
            messagebox.showerror("Error", "El archivo CRM no contiene una columna 'Serie'.")
            return
        series_crm = crm_df["Serie"].astype(str).unique().tolist()
    except Exception as e:
        messagebox.showerror("Error", f"No se pudo leer el archivo CRM:\n{e}")
        return

    fallas, alertas = [], []
    for item in data:
        serie = str(item.get("Serie", ""))
        falla_alerta = item.get("Estatus", "").lower()

        if serie not in series_crm:
            if "falla" in falla_alerta:
                fallas.append(item)
            elif "alerta" in falla_alerta or "demanda" in falla_alerta:
                alertas.append(item)

    # === 6️⃣ Seleccionar carpeta para guardar ===
    save_dir = filedialog.askdirectory(title="Selecciona carpeta para guardar los resultados")
    if not save_dir:
        messagebox.showinfo("Cancelado", "No seleccionaste carpeta. Operación cancelada.")
        return

    # === 7️⃣ Guardar los resultados ===
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
        messagebox.showinfo("Sin nuevos registros", "No se encontraron fallas o alertas fuera del CRM.")

if __name__ == "__main__":
    main()
