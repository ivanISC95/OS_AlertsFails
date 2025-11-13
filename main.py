DEBUG = True  # Cambia a False cuando no quieras ver los prints

def debug_print(*args):
    if DEBUG:
        print(*args)

import requests
import pandas as pd
import tkinter as tk
from tkinter import filedialog, simpledialog, messagebox, ttk
import pickle
import sys
import os



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
def resource_path(relative_path):
    """Obtiene la ruta correcta del recurso tanto en desarrollo como en el .exe"""
    try:
        # Cuando el script está empacado con PyInstaller
        base_path = sys._MEIPASS
    except Exception:
        # Cuando se ejecuta como script normal
        base_path = os.path.abspath(".")

    return os.path.join(base_path, relative_path)

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
        login_response = requests.post(login_url, json=login_body, verify=False)
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

            crm_df["Fecha de creación"] = pd.to_datetime(
        # Convertir fechas y filtrar por mes más antiguo
        if "Fecha de creación" in crm_df.columns:
                crm_df["Fecha de creación"],
                format="%d/%m/%Y  %I:%M:%S %p",
                errors="coerce"
            )
            fecha_mayor = crm_df["Fecha de creación"].max()
            if pd.notna(fecha_mayor):
                mes_mayor = fecha_mayor.month
                # Guardamos las series que pertenecen al mes más antiguo
                series_mes_mayor = crm_df.loc[
                    crm_df["Fecha de creación"].dt.month == mes_mayor, "Serie"
                ].astype(str).unique().tolist()
            else:
                series_mes_mayor = crm_df["Serie"].astype(str).unique().tolist()
        else:
            series_mes_mayor = crm_df["Serie"].astype(str).unique().tolist()

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
        control = item.get("ControlDeActivos", "").strip().lower()
        estatus = item.get("EstatusKOF", "").strip().upper()
        direccion_campos = [
            item.get("CodigoPostal", "").strip(),
            item.get("EntreCalles", "").strip(),
            item.get("DirecciónPdV", "").strip(),
            item.get("PdV", "").strip(),
            item.get("IdPdV", "").strip()
        ]

        # ❌ Descartar si cumple cualquiera de las condiciones:
        # 1. ControlDeActivos = 'Sin coincidencia'
        # 2. EstatusKOF = 'LEGL'
        # 3. Cualquier campo de dirección está vacío
        if (
                control == "sin coincidencia"
                or estatus == "LEGL"
                or any(campo == "" for campo in direccion_campos)
        ):
            return False

        return True

    fallas, alertas = [], []
    debug_print(f"Total alertas/fallas obtenidas de API: {len(data)}")
    debug_print(f"Series CRM: {len(series_crm)}, Series mes mayor: {len(series_mes_mayor)}")
    debug_print(f"Series excluidas de vaultlist: {len(series_excluir)}")

    for item in data:
        serie = str(item.get("Serie", ""))
        falla_alerta = item.get("Estatus", "").lower()
        region = item.get("Region", "").strip().lower()
        falla_tipo_alerta = item.get("FallaAlerta", "").lower()

        # Saltar si no pasa los filtros básicos
        if serie in series_crm:
            debug_print(f"[IGNORADA - CRM] Serie {serie}")
            continue
        if serie in series_excluir:
            debug_print(f"[IGNORADA - Vaultlist] Serie {serie}")
            continue
        if not pasa_filtros(item):
            debug_print(
                f"[IGNORADA - Filtros] Serie {serie} (ControlDeActivos={item.get('ControlDeActivos')}, EstatusKOF={item.get('EstatusKOF')})")
            continue

        # Si llega aquí, pasa filtros básicos
        if ("falla" in falla_alerta or "falla" in falla_tipo_alerta or "temperatura" in falla_tipo_alerta):
            if serie not in series_mes_mayor:
                fallas.append(item)
                debug_print(f"[FALLA ✅] Serie {serie}")
            else:
                debug_print(f"[IGNORADA - Mes mayor] Serie {serie}")
        elif "alta demanda de compresor" in falla_tipo_alerta and region == 'monarca':
            alertas.append(item)
            debug_print(f"[ALERTA ✅] Serie {serie} (Región={region})")
        else:
            debug_print(
                f"[IGNORADA - No cumple condición de alerta/falla] Serie {serie} (Estatus={falla_alerta}, Región={region})")

    debug_print(f"Total fallas válidas: {len(fallas)}")
    debug_print(f"Total alertas válidas: {len(alertas)}")

    progreso.destroy()

    # === 8️⃣ Guardar archivos ===
    save_dir = filedialog.askdirectory(title="Selecciona carpeta para guardar resultados")
    if not save_dir:
        messagebox.showinfo("Cancelado", "No seleccionaste carpeta.")
        return

    def formatear_dataframe(data):
        df = pd.DataFrame(data)
        # === Crear columnas base ===
        df["RAZON SOCIAL"] = df.get("PdV", "")
        df["RESPONSABLE"] = df.get("ContactoDePdV", "")
        df["MODELO"] = df.get("Modelo", "")
        df["No SERIE"] = df.get("Serie", "")
        df["FALLA"] = df.get("FallaAlerta", "")
        df["REPORTE IMAGEN"] = ""
        df["CALLE Y NUMERO"] = df.get("DirecciónPdV", "")
        df["ENTRE CALLE"] = df.get("EntreCalles", "")
        df["Y CALLE"] = ""

        # --- Separar ENTRE CALLE y Y CALLE ---
        def separar_calles(valor):
            if isinstance(valor, str) and "y" in valor.lower():
                partes = valor.split("y", 1)
                entre = partes[0].strip()
                ycalle = partes[1].strip()
                return entre, ycalle
            return valor, ""

        df["ENTRE CALLE"], df["Y CALLE"] = zip(*df["ENTRE CALLE"].map(separar_calles))

        # === Cargar CP cp_data.pkl ===
        pkl_path = resource_path("cp_data.pkl")

        if not os.path.exists(pkl_path):
            raise FileNotFoundError(f"No se encontró el archivo: {pkl_path}")

        with open(pkl_path, "rb") as f:
            cp_data = pickle.load(f)
        df_cp = pd.DataFrame(cp_data)

        # Normalizar tipos y columnas
        df_cp.rename(columns={
            "d_codigo": "CodigoPostal",
            "d_asenta": "COLONIA / POBLADO",
            "D_mnpio": "DELEGACION / MUNICIPIO / CIUDAD"
        }, inplace=True)

        # Convertir CodigoPostal a string (para evitar errores al comparar)
        df["CodigoPostal"] = df.get("CodigoPostal", "").astype(str)
        df_cp["CodigoPostal"] = df_cp["CodigoPostal"].astype(str)

        # --- Unir la información del JSON con el DataFrame principal ---
        df = df.merge(
            df_cp[["CodigoPostal", "COLONIA / POBLADO", "DELEGACION / MUNICIPIO / CIUDAD"]],
            on="CodigoPostal",
            how="left"
        )

        # Código postal, observaciones, coordenadas, etc.
        df["CP"] = df.get("CodigoPostal", "")
        df["NUM. TEL"] = df.get("NumeroTelefono", "")
        df["HORARIO DE ATENCION"] = ""
        df["No CLIENTE DETALLISTA"] = df.get("IdPdV", "")
        df["CEDIS/DISTRIBUIDORA"] = ""
        df["OBSERVACIONES"] = "enfriador reportado por conectividad"
        df["SOLICITUD DE SERVICIO"] = ""
        df["LON"] = df.get("UltimaLongitud", "")
        df["LAT"] = df.get("UltimaLatitud", "")
        df["ID REPORTE/TICKET/FOLIO/PAEEEM"] = ""
        df["OS"] = ""
        # === Reordenar columnas finales ===
        columnas_finales = [
            "RAZON SOCIAL",
            "RESPONSABLE",
            "MODELO",
            "No SERIE",
            "FALLA",
            "REPORTE IMAGEN",
            "CALLE Y NUMERO",
            "ENTRE CALLE",
            "Y CALLE",
            "COLONIA / POBLADO",
            "DELEGACION / MUNICIPIO / CIUDAD",
            "CP",
            "NUM. TEL",
            "HORARIO DE ATENCION",
            "No CLIENTE DETALLISTA",
            "CEDIS/DISTRIBUIDORA",
            "OBSERVACIONES",
            "SOLICITUD DE SERVICIO",
            "LON",
            "LAT",
            "ID REPORTE/TICKET/FOLIO/PAEEEM",
            "OS"
        ]

        df_final = df[columnas_finales]

        return df_final

    saved_files = []
    debug_print(f"ALERTAS QUE VAN AL EXCEL: {len(alertas)}")
    if fallas:
        df_fallas = formatear_dataframe(fallas)

        # 🔹 Eliminar duplicados por la columna "Serie"
        if "No SERIE" in df_fallas.columns:
            df_fallas = df_fallas.drop_duplicates(subset="No SERIE", keep="first")

        fallas_path = os.path.join(save_dir, "Fallas_Nuevas.xlsx")
        df_fallas.to_excel(fallas_path, index=False)
        saved_files.append(fallas_path)

    if alertas:
        df_alertas = formatear_dataframe(alertas)

        # 🔹 Eliminar duplicados por la columna "Serie"
        if "No SERIE" in df_alertas.columns:
            df_alertas = df_alertas.drop_duplicates(subset="No SERIE", keep="first")

        alertas_path = os.path.join(save_dir, "Alertas_Nuevas.xlsx")
        df_alertas.to_excel(alertas_path, index=False)
        saved_files.append(alertas_path)

    if saved_files:
        messagebox.showinfo("Completado", "✅ Archivos generados:\n\n" + "\n".join(saved_files))
    else:
        messagebox.showinfo("Sin nuevos registros", "No se encontraron alertas o fallas fuera del CRM.")


if __name__ == "__main__":
    main()
