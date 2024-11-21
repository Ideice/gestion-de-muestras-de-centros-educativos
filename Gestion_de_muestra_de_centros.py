import tkinter as tk
from tkinter import filedialog, messagebox, ttk
import pandas as pd
import random
from math import ceil
from scipy.stats import norm
import folium
from folium.plugins import MarkerCluster
import webbrowser
import io
import numpy as np
from sklearn.cluster import KMeans  # Para agrupar las ubicaciones


# Inicializar la ventana principal
root = tk.Tk()
root.title("Gestión de Muestras - Evaluación e Investigación Educativa")
#root.geometry("1400x900")  # Aumentar el tamaño de la ventana 1200x800
root.geometry("1900x1000")  # Aumentar el tamaño de la ventana 1200x800
root.configure(bg="#ADD8E6")

# Variables globales
df_poblacion = pd.DataFrame()  # Para almacenar los datos del archivo Excel
df_muestra = pd.DataFrame()  # Para almacenar la muestra seleccionada
tamano_muestra = 0
tipo_muestra = tk.StringVar()
numero_grupos = tk.IntVar(value=1)  # Variable para almacenar el número de grupos

##Incluí esta funcion para ver tamaño de pantalla
# Función para mostrar el tamaño de la ventana en tiempo real (opcional)
def actualizar_tamano_ventana(event=None):
    tamano = f"Tamaño de ventana: {root.winfo_width()}x{root.winfo_height()}"
    lbl_tamano_ventana.config(text=tamano)

# Etiqueta para mostrar el tamaño de la ventana en la esquina superior derecha
lbl_tamano_ventana = tk.Label(root, text="", bg="#ADD8E6", fg="black", font=("Arial", 10))
lbl_tamano_ventana.grid(row=0, column=1, sticky="ne", padx=10, pady=10)

# Llamar a la función para actualizar el tamaño de la ventana cuando se redimensione
root.bind("<Configure>", actualizar_tamano_ventana)

##

# Función para cargar el archivo Excel
def cargar_datos_excel():
    global df_poblacion
    try:
        file_path = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx")])
        if file_path:
            df_poblacion = pd.read_excel(file_path)
            lbl_nombre_archivo.config(text=f"{'Filename:'} {file_path.split("/")[-1]}", fg="blue")
            lbl_cantidad_registros.config(
                text=f"{'Cantidad de registros:'} {len(df_poblacion)}"
            )
            # lbl_tamano_muestra.config(text=f"{'Tamaño muestra: '} {tamano_muestra}")
            lbl_cantidad_campos.config(text=f"{'Cantidad de campos:'} {len(df_poblacion.columns)}", fg="blue")
            messagebox.showinfo(
                "Éxito", "Datos cargados correctamente desde el archivo Excel."
            )
        else:
            messagebox.showwarning(
                "Advertencia", "No se ha seleccionado ningún archivo."
            )
    except Exception as e:
        messagebox.showerror("Error", f"No se pudo cargar el archivo Excel: {e}")



# Función para calcular el tamaño de la muestra
def calcular_tamano_muestra():
    global tamano_muestra
    try:
        if df_poblacion.empty:
            messagebox.showwarning("Advertencia", "Cargue primero un archivo.")
            return
      

        margen_error = float(entry_margen_error.get()) / 100
        nivel_confianza = float(entry_nivel_confianza.get()) / 100
        '''
        # Crear el campo de entrada para el margen de error
        entry_margen_error = tk.Entry(frame_confianza_error, font=("Arial", 14, "bold"), justify="center")
        entry_margen_error.grid(row=3, column=0, padx=5, pady=5)

        # Crear el campo de entrada para el nivel de confianza
        entry_nivel_confianza = tk.Entry(frame_confianza_error, font=("Arial", 14, "bold"), justify="center")
        entry_nivel_confianza.grid(row=1, column=0, padx=5, pady=5)
        '''
        # Obtener el valor Z correspondiente al nivel de confianza
        z_score = norm.ppf(1 - (1 - nivel_confianza) / 2)
        p = 0.5  # Proporción estimada (máxima variabilidad)
        N = len(df_poblacion)

        # Fórmula del tamaño de muestra para población finita
        n_0 = (z_score**2 * p * (1 - p)) / (margen_error**2)
        n = n_0 / (1 + (n_0 - 1) / N)
        tamano_muestra = ceil(n)

        lbl_tamano_muestra.config(text=f"{'Tamaño muestra: '} {tamano_muestra}", font=("Arial", 11, "bold"), anchor="center")

    except ValueError:
        messagebox.showerror(
            "Error",
            "Ingrese valores válidos para el nivel de confianza y margen de error.",
        )
    except Exception as e:
        messagebox.showerror("Error", f"Error en el cálculo del tamaño de muestra: {e}")


# Función para seleccionar el tipo de muestra y generar el resultado
def seleccionar_tipo_muestra():
    tipo = tipo_muestra.get()

    if tipo == "Aleatoria Simple":
        seleccionar_muestra_aleatoria_simple()
    elif tipo == "Aleatoria Simple Sistemática":
        seleccionar_muestra_sistematica()
    elif "Estratificada por" in tipo:
        # Desglosar el tipo de estratificación
        campo_estratificacion = tipo.split(" por ")[-1].lower().replace(" ", "_")
        seleccionar_muestra_estratificada_por_campo(campo_estratificacion)
    elif "Estratificada Sistemática por" in tipo:
        # Desglosar el tipo de estratificación sistemática
        campo_estratificacion = tipo.split(" por ")[-1].lower().replace(" ", "_")
        seleccionar_muestra_estratificada_sistematica(campo_estratificacion)
    else:
        messagebox.showwarning("Advertencia", "Seleccione un tipo de muestra válido.")


# Función para seleccionar muestra aleatoria simple
def seleccionar_muestra_aleatoria_simple():
    global df_muestra
    try:
        if df_poblacion.empty or tamano_muestra == 0:
            messagebox.showwarning(
                "Advertencia",
                "Asegúrese de cargar un archivo y calcular el tamaño de la muestra.",
            )
            return

        df_muestra = df_poblacion.sample(n=tamano_muestra, random_state=42)
        mostrar_muestra(df_muestra)

    except Exception as e:
        messagebox.showerror("Error", f"Error al seleccionar la muestra: {e}")


# Función para seleccionar muestra sistemática con un primer número aleatorio y vuelta circular
def seleccionar_muestra_sistematica():
    global df_muestra
    try:
        if df_poblacion.empty or tamano_muestra == 0:
            messagebox.showwarning(
                "Advertencia",
                "Asegúrese de cargar un archivo y calcular el tamaño de la muestra.",
            )
            return

        # Tamaño de la población
        tamano_poblacion = len(df_poblacion)

        # Calcular el intervalo
        intervalo = tamano_poblacion // tamano_muestra

        # Seleccionar el primer elemento de forma aleatoria entre 0 y el tamaño total de la población - 1
        primer_elemento = random.randint(0, tamano_poblacion - 1)

        # Lista para almacenar los índices seleccionados
        indices_seleccionados = []

        # Realizar la selección sistemática
        for i in range(tamano_muestra):
            # Calcular el índice del elemento actual (modular para circularidad)
            indice_actual = (primer_elemento + i * intervalo) % tamano_poblacion
            indices_seleccionados.append(indice_actual)

        # Seleccionar los elementos correspondientes a los índices en la población
        df_muestra = df_poblacion.iloc[indices_seleccionados]

        # Mostrar la muestra seleccionada en la tabla
        mostrar_muestra(df_muestra)

    except Exception as e:
        messagebox.showerror(
            "Error", f"Error al seleccionar la muestra sistemática: {e}"
        )

# Función ajustada para seleccionar muestra estratificada por el campo seleccionado
def seleccionar_muestra_estratificada_por_campo(campo_estratificacion):
    global df_muestra
    try:
        if df_poblacion.empty or tamano_muestra == 0:
            messagebox.showwarning(
                "Advertencia",
                "Asegúrese de cargar un archivo y calcular el tamaño de la muestra.",
            )
            return

        if campo_estratificacion not in df_poblacion.columns:
            messagebox.showerror(
                "Error", f"El campo '{campo_estratificacion}' no existe en los datos."
            )
            return

        # Agrupar la población por el campo de estratificación seleccionado
        grupos = df_poblacion.groupby(campo_estratificacion)

        # Crear un DataFrame vacío para almacenar la muestra final
        df_muestra = pd.DataFrame()

        # Selección proporcional por cada grupo (estrato)
        tamanos_estratos = []
        for nombre_estrato, grupo in grupos:
            # Calcular la proporción del grupo en la población
            proporcion = len(grupo) / len(df_poblacion)

            # Calcular el tamaño de muestra para este estrato (redondeo más equitativo)
            tamano_estrato = max(1, round(tamano_muestra * proporcion))  # Cambiar a round()
            tamanos_estratos.append(tamano_estrato)

            # Seleccionar aleatoriamente la muestra del estrato
            muestra_estrato = grupo.sample(n=tamano_estrato, random_state=42)
            df_muestra = pd.concat([df_muestra, muestra_estrato])

        # Verificar si el número total de la muestra es mayor o menor al tamaño de muestra esperado
        total_seleccionado = len(df_muestra)
        diferencia = tamano_muestra - total_seleccionado

        if diferencia > 0:
            # Si la muestra es menor, seleccionamos elementos adicionales aleatoriamente de todos los estratos
            grupos_faltantes = [grupo for nombre_estrato, grupo in grupos if len(grupo) > 0]
            while diferencia > 0:
                for grupo in grupos_faltantes:
                    if diferencia == 0:
                        break
                    # Seleccionar 1 elemento adicional por iteración
                    muestra_adicional = grupo.sample(n=1, random_state=42)
                    df_muestra = pd.concat([df_muestra, muestra_adicional])
                    diferencia -= 1

        elif diferencia < 0:
            # Si seleccionamos más de la muestra esperada, reducimos de la muestra total
            df_muestra = df_muestra.sample(n=tamano_muestra, random_state=42)

        # Mostrar la muestra seleccionada en la tabla
        mostrar_muestra(df_muestra)

    except Exception as e:
        messagebox.showerror(
            "Error", f"Error al seleccionar la muestra estratificada por {campo_estratificacion}: {e}"
        )

# Función para seleccionar muestra estratificada sistemática por el campo seleccionado
def seleccionar_muestra_estratificada_sistematica(campo_estratificacion):
    global df_muestra
    try:
        if df_poblacion.empty or tamano_muestra == 0:
            messagebox.showwarning(
                "Advertencia",
                "Asegúrese de cargar un archivo y calcular el tamaño de la muestra.",
            )
            return

        if campo_estratificacion not in df_poblacion.columns:
            messagebox.showerror(
                "Error", f"El campo '{campo_estratificacion}' no existe en los datos."
            )
            return

        # Agrupar la población por el campo de estratificación seleccionado
        grupos = df_poblacion.groupby(campo_estratificacion)

        # Crear un DataFrame vacío para almacenar la muestra final
        df_muestra = pd.DataFrame()

        # Selección sistemática proporcional por cada grupo (estrato)
        tamanos_estratos = []
        for nombre_estrato, grupo in grupos:
            # Calcular la proporción del grupo en la población
            proporcion = len(grupo) / len(df_poblacion)

            # Calcular el tamaño de muestra para este estrato (redondeo más equitativo)
            tamano_estrato = max(1, round(tamano_muestra * proporcion))
            tamanos_estratos.append(tamano_estrato)

            # Selección sistemática dentro del estrato
            if len(grupo) > tamano_estrato:
                intervalo = len(grupo) // tamano_estrato

                # Seleccionar el primer elemento de forma aleatoria
                primer_elemento = random.randint(0, len(grupo) - 1)

                # Lista para almacenar los índices seleccionados dentro del grupo (estrato)
                indices_seleccionados = []

                for i in range(tamano_estrato):
                    indice_actual = (primer_elemento + i * intervalo) % len(grupo)
                    indices_seleccionados.append(grupo.index[indice_actual])

                # Agregar las muestras del estrato al DataFrame final
                muestra_estrato = grupo.loc[indices_seleccionados]
                df_muestra = pd.concat([df_muestra, muestra_estrato])
            else:
                # Si el estrato tiene menos registros que el tamaño de muestra requerido, se toma todo el estrato
                df_muestra = pd.concat([df_muestra, grupo])

        # Mostrar la muestra seleccionada en la tabla
        mostrar_muestra(df_muestra)

    except Exception as e:
        messagebox.showerror(
            "Error", f"Error al seleccionar la muestra estratificada sistemática por {campo_estratificacion}: {e}"
        )

# Ejemplo de cómo usar la función con diferentes campos
def seleccionar_muestra_por_regional():
    seleccionar_muestra_estratificada_por_campo('regional')

def seleccionar_muestra_por_sector():
    seleccionar_muestra_estratificada_por_campo('sector')

def seleccionar_muestra_por_zona():
    seleccionar_muestra_estratificada_por_campo('zona')

def seleccionar_muestra_por_nivel_educativo():
    seleccionar_muestra_estratificada_por_campo('nivel_educativo')

 
##
def generar_mapa_centros():
    global df_muestra
    if df_muestra.empty:
        messagebox.showwarning("Advertencia", "No hay muestra seleccionada para generar el mapa.")
        return

    try:
        # Convertir las columnas de latitud y longitud a numéricas, ignorando los errores
        df_muestra['latitud'] = pd.to_numeric(df_muestra['latitud'], errors='coerce')
        df_muestra['longitud'] = pd.to_numeric(df_muestra['longitud'], errors='coerce')

        # Eliminar filas con valores NaN en latitud o longitud
        df_muestra = df_muestra.dropna(subset=['latitud', 'longitud'])

        # Calcular la latitud y longitud media para centrar el mapa
        latitud_media = df_muestra['latitud'].mean()
        longitud_media = df_muestra['longitud'].mean()

        # Crear un mapa centrado en la latitud y longitud medias
        mapa = folium.Map(location=[latitud_media, longitud_media], zoom_start=7)

        # Crear un cluster de marcadores
        marker_cluster = MarkerCluster().add_to(mapa)

        # Añadir cada centro en el cluster de marcadores
        for _, row in df_muestra.iterrows():
            folium.Marker(
                location=[row['latitud'], row['longitud']],
                popup=f"{row['centro']} ({row['regional']})",
                icon=folium.Icon(color='blue')
            ).add_to(marker_cluster)

        # Guardar el mapa en un archivo HTML
        mapa.save('mapa_centros.html')
        messagebox.showinfo("Éxito", "Mapa generado en formato HTML.")
        import webbrowser
        webbrowser.open('mapa_centros.html')

    except Exception as e:
        messagebox.showerror("Error", f"Error al generar el mapa de centros: {e}")

# Función para mostrar la muestra en pantalla con barra deslizadora
def mostrar_muestra(muestra):
    tree_muestra.delete(*tree_muestra.get_children())  # Limpiar la tabla anterior

    for i, (index, row) in enumerate(muestra.iterrows(), start=1):
        tree_muestra.insert(
            "",
            "end",
            values=(
                i,
                row["num"],
                row["regional"],
                row["centro"],
                row["sector"],
                row["zona"],
                row["nivel_educativo"],
                row["latitud"],
                row["longitud"],
            ),
        )

# Función para exportar la muestra seleccionada
def exportar_muestra():
    if df_muestra.empty:
        messagebox.showwarning(
            "Advertencia", "No hay muestra seleccionada para exportar."
        )
        return

    try:
        file_path = filedialog.asksaveasfilename(
            defaultextension=".xlsx", filetypes=[("Excel files", "*.xlsx")]
        )
        if file_path:
            df_muestra.to_excel(file_path, index=False)
            messagebox.showinfo("Éxito", "Muestra exportada a Excel.")
    except Exception as e:
        messagebox.showerror("Error", f"Error al exportar la muestra: {e}")


# Sección 1: Cargar archivo Excel
frame_archivo = tk.Frame(root, bg="#ADD8E6", relief="solid", bd=3)
#frame_archivo.grid(row=0, column=0, padx=10, pady=10)
frame_archivo.grid(row=0, column=0, padx=0, pady=0)

btn_cargar_archivo = tk.Button(
    frame_archivo,
    text="Cargar archivo para extraer muestra",
    command=cargar_datos_excel,
    width=40,
)
btn_cargar_archivo.grid(row=0, column=0, padx=5, pady=5)


lbl_nombre_archivo = tk.Label(frame_archivo, text="Nombre de archivo: ", bg="#ADD8E6")
lbl_nombre_archivo.grid(row=1, column=0, padx=5, pady=5)

lbl_cantidad_registros = tk.Label(
    frame_archivo, text="Cantidad de registros: ", bg="#ADD8E6"
)
lbl_cantidad_registros.grid(row=2, column=0, padx=5, pady=5)

lbl_cantidad_campos = tk.Label(frame_archivo, text="Cantidad de campos: ", bg="#ADD8E6")
lbl_cantidad_campos.grid(row=3, column=0, padx=5, pady=5)

# Sección 2: Nivel de confianza y margen de error
frame_confianza_error = tk.Frame(root, bg="#ADD8E6", relief="solid", bd=2)
frame_confianza_error.grid(row=1, column=0, padx=10, pady=10)

lbl_confianza = tk.Label(
    frame_confianza_error, text="Nivel de confianza (%)", bg="#ADD8E6"
)
lbl_confianza.grid(row=0, column=0, padx=5, pady=5)

entry_nivel_confianza = tk.Entry(frame_confianza_error)
entry_nivel_confianza.grid(row=1, column=0, padx=5, pady=5)

lbl_margen_error = tk.Label(
    frame_confianza_error, text="Margen de error (%)", bg="#ADD8E6"
)
lbl_margen_error.grid(row=2, column=0, padx=5, pady=5)

entry_margen_error = tk.Entry(frame_confianza_error)
entry_margen_error.grid(row=3, column=0, padx=5, pady=5)

# Botón para calcular tamaño de muestra
btn_calcular_muestra = tk.Button(
    frame_confianza_error,
    text="Calcular Tamaño de Muestra",
    command=calcular_tamano_muestra,
)
btn_calcular_muestra.grid(row=4, column=0, padx=5, pady=5)

lbl_tamano_muestra = tk.Label(
    frame_confianza_error, text="Tamaño de muestra: ", bg="#ADD8E6"
)
lbl_tamano_muestra.grid(row=5, column=0, padx=5, pady=5)

# Sección 3: Selección del tipo de muestra
frame_tipo_muestra = tk.Frame(root, bg="#ADD8E6", relief="solid", bd=2)
frame_tipo_muestra.grid(row=2, column=0, padx=10, pady=10)

lbl_tipo_muestra = tk.Label(
    frame_tipo_muestra, text="Seleccione el tipo de muestra", bg="#ADD8E6"
)
lbl_tipo_muestra.grid(row=0, column=0, padx=5, pady=5)


# Ampliar el ancho del combobox a un valor que acomode las etiquetas largas
combo_tipo_muestra = ttk.Combobox(
    frame_tipo_muestra, textvariable=tipo_muestra, width=40
)  # Cambiar el valor de width
combo_tipo_muestra["values"] = [
    "Aleatoria Simple",
    "Aleatoria Simple Sistemática",
    "Estratificada por Regional",
    "Estratificada por Zona",
    "Estratificada por Sector",
    "Estratificada por Nivel Educativo",
    "Estratificada Sistemática por Regional",
    "Estratificada Sistemática por Zona",
    "Estratificada Sistemática por Sector",
    "Estratificada Sistemática por Nivel Educativo",
]
combo_tipo_muestra.grid(row=1, column=0, padx=5, pady=5)


btn_seleccionar_muestra = tk.Button(
    frame_tipo_muestra, text="Generar Muestra", command=seleccionar_tipo_muestra
)
btn_seleccionar_muestra.grid(row=2, column=0, padx=5, pady=5)

# Sección 4: Visualización de la muestra con barra deslizadora
frame_resultado = tk.Frame(root, bg="#ADD8E6", relief="solid", bd=2)
frame_resultado.grid(row=3, column=0, padx=10, pady=10)

# Crear una barra deslizadora para la tabla
frame_tree_scroll = tk.Frame(frame_resultado)
frame_tree_scroll.pack(fill="both", expand=True)

scrollbar_y = tk.Scrollbar(frame_tree_scroll, orient="vertical")
scrollbar_y.pack(side="right", fill="y")

scrollbar_x = tk.Scrollbar(frame_tree_scroll, orient="horizontal")
scrollbar_x.pack(side="bottom", fill="x")

# Crear un estilo para personalizar los encabezados del Treeview
style = ttk.Style()
style.configure(
    "Treeview.Heading", font=("Arial", 10, "bold"), foreground="blue"
)  # Encabezados en azul y negrita

tree_muestra = ttk.Treeview(
    frame_tree_scroll,
    columns=(
        "Elemento",
        "num",
        "regional",
        "centro",
        "sector",
        "zona",
        "nivel_educativo",
        "latitud",
        "longitud",
    ),
    show="headings",
    yscrollcommand=scrollbar_y.set,
    xscrollcommand=scrollbar_x.set,
)
tree_muestra.heading("Elemento", text="Elemento")
tree_muestra.heading("num", text="Num")
tree_muestra.heading("regional", text="Regional")
tree_muestra.heading("centro", text="Centro")
tree_muestra.heading("sector", text="Sector")
tree_muestra.heading("zona", text="Zona")
tree_muestra.heading("nivel_educativo", text="Nivel educativo")
tree_muestra.heading("latitud", text="Latitud")
tree_muestra.heading("longitud", text="Longitud")

tree_muestra.column("Elemento", anchor="center", width=80)
tree_muestra.column("num", anchor="center", width=70)
tree_muestra.column("regional", anchor="center", width=200)
tree_muestra.column("centro", anchor="center", width=425)
tree_muestra.column("sector", anchor="center", width=100)
tree_muestra.column("zona", anchor="center", width=100)
tree_muestra.column("nivel_educativo", anchor="center", width=180)
tree_muestra.column("latitud", anchor="center", width=120)
tree_muestra.column("longitud", anchor="center", width=120)

# Configuración de los encabezados en negrita
style = ttk.Style()
style.configure("Treeview.Heading", font=("Arial", 10, "bold"))

tree_muestra.pack(fill="both", expand=True)

scrollbar_y.config(command=tree_muestra.yview)
scrollbar_x.config(command=tree_muestra.xview)

# Sección 5: Exportación de la muestra seleccionada y visualización de centros
frame_exportar = tk.Frame(root, bg="#ADD8E6", relief="solid", bd=2)
frame_exportar.grid(row=4, column=0, padx=10, pady=10)

# Botón para exportar la muestra a Excel
btn_exportar_excel = tk.Button(
    frame_exportar, text="Exportar a Excel", command=exportar_muestra
)
btn_exportar_excel.grid(row=0, column=0, padx=5, pady=5)

# Botón para generar la visualización del mapa
btn_generar_mapa = tk.Button(
    frame_exportar, text="Generar Visualización de Centros", command=generar_mapa_centros
)
btn_generar_mapa.grid(row=1, column=0, padx=5, pady=5)

btn_quit = tk.Button(
    frame_exportar, text="QUIT", command=root.quit
)
btn_quit.grid(row=3, column=0, padx=5, pady=5)


# ----------- NUEVA SECCIÓN PARA CREACIÓN DE RUTAS -----------

# Función para agrupar los centros seleccionados en grupos usando k-means clustering
def crear_rutas():
    global df_muestra
    if df_muestra.empty:
        messagebox.showwarning("Advertencia", "No hay muestra seleccionada.")
        return

    try:
        # Convertir las columnas de latitud y longitud a numéricas
        df_muestra['latitud'] = pd.to_numeric(df_muestra['latitud'], errors='coerce')
        df_muestra['longitud'] = pd.to_numeric(df_muestra['longitud'], errors='coerce')

        # Eliminar filas con NaN en latitud o longitud
        df_muestra = df_muestra.dropna(subset=['latitud', 'longitud'])

        if df_muestra.empty:
            messagebox.showerror("Error", "No hay suficientes datos de ubicación para crear grupos.")
            return

        # Obtener el número de grupos
        num_grupos = numero_grupos.get()
        if num_grupos < 1:
            messagebox.showwarning("Advertencia", "El número de grupos debe ser mayor que 0.")
            return

        # Extraer coordenadas
        coordenadas = df_muestra[['latitud', 'longitud']].values

        # Agrupar usando k-means
        kmeans = KMeans(n_clusters=num_grupos, random_state=0).fit(coordenadas)
        df_muestra['grupo'] = kmeans.labels_

        # Calcular la ruta óptima dentro de cada grupo
        df_muestra['orden_visita'] = 0  # Columna para el orden de visita

        for grupo in range(num_grupos):
            subgrupo = df_muestra[df_muestra['grupo'] == grupo]
            indices = subgrupo.index.tolist()
            ruta_ordenada = calcular_ruta_optima(subgrupo[['latitud', 'longitud']].values)
            for orden, idx in enumerate(ruta_ordenada, start=1):
                df_muestra.loc[indices[idx], 'orden_visita'] = orden

        messagebox.showinfo("Éxito", f"Rutas creadas para {num_grupos} grupos con orden de visita asignado.")
        visualizar_rutas()

    except Exception as e:
        messagebox.showerror("Error", f"Error al crear grupos/ruta: {e}")


def calcular_ruta_optima(coordenadas):
    """
    Calcula la ruta óptima (heurística) usando el vecino más cercano.
    """
    n = len(coordenadas)
    visitado = [False] * n
    ruta = []
    actual = 0
    visitado[actual] = True
    ruta.append(actual)

    for _ in range(1, n):
        distancia_minima = float('inf')
        siguiente = None

        for i in range(n):
            if not visitado[i]:
                distancia = np.linalg.norm(coordenadas[actual] - coordenadas[i])
                if distancia < distancia_minima:
                    distancia_minima = distancia
                    siguiente = i

        if siguiente is not None:
            visitado[siguiente] = True
            ruta.append(siguiente)
            actual = siguiente

    return ruta


def visualizar_rutas():
    global df_muestra
    try:
        # Crear el mapa
        latitud_media = df_muestra['latitud'].mean()
        longitud_media = df_muestra['longitud'].mean()
        mapa = folium.Map(location=[latitud_media, longitud_media], zoom_start=7)

        # Crear colores únicos para cada grupo
        colores_grupos = ['red', 'blue', 'green', 'purple', 'orange', 'darkred', 'lightred', 'darkblue', 'darkgreen']

        for grupo in df_muestra['grupo'].unique():
            subgrupo = df_muestra[df_muestra['grupo'] == grupo]
            color = colores_grupos[grupo % len(colores_grupos)]

            # Añadir marcadores y rutas al mapa
            puntos = []
            for _, row in subgrupo.sort_values('orden_visita').iterrows():
                folium.Marker(
                    location=[row['latitud'], row['longitud']],
                    popup=f"Centro: {row['centro']}<br>Grupo: {grupo + 1}<br>Orden: {row['orden_visita']}",
                    icon=folium.Icon(color=color)
                ).add_to(mapa)
                puntos.append([row['latitud'], row['longitud']])

            # Dibujar la línea de la ruta
            if len(puntos) > 1:
                folium.PolyLine(puntos, color=color, weight=2.5, opacity=1).add_to(mapa)

        # Guardar el mapa
        mapa.save("rutas_centros.html")
        webbrowser.open("rutas_centros.html")

    except Exception as e:
        messagebox.showerror("Error", f"Error al visualizar las rutas: {e}")


# ----------- INTERFAZ GRÁFICA -----------

# Nuevo frame para el cálculo de rutas
frame_rutas = tk.Frame(root, bg="#ADD8E6", relief="solid", bd=2)
frame_rutas.grid(row=2, column=1, padx=10, pady=10)

# Etiqueta para el número de grupos
lbl_numero_grupos = tk.Label(frame_rutas, text="Número de grupos:", bg="#ADD8E6")
lbl_numero_grupos.grid(row=0, column=0, padx=5, pady=5)

# Campo de entrada para el número de grupos
entry_numero_grupos = tk.Entry(frame_rutas, textvariable=numero_grupos)
entry_numero_grupos.grid(row=0, column=1, padx=5, pady=5)

# Botón para crear rutas
btn_crear_rutas = tk.Button(frame_rutas, text="Crear Grupos/Rutas", command=crear_rutas)
btn_crear_rutas.grid(row=1, column=0, columnspan=2, padx=5, pady=5)

# Botón para visualizar rutas en el mapa
#btn_visualizar_rutas = tk.Button(frame_rutas, text="Crear/visualizar Rutas", command=visualizar_rutas)
#btn_visualizar_rutas.grid(row=2, column=0, columnspan=2, padx=5, pady=5)

# Función para generar la tabla con los grupos creados
def generar_tabla_grupos():
    global df_muestra
    if df_muestra.empty:
        messagebox.showwarning("Advertencia", "No hay grupos creados para mostrar.")
        return

    try:
        # Limpiar la tabla anterior
        tree_grupos.delete(*tree_grupos.get_children())

        # Llenar la tabla con los datos
        for _, row in df_muestra.sort_values(['grupo', 'orden_visita']).iterrows():
            tree_grupos.insert(
                "",
                "end",
                values=(
                    row['grupo'] + 1,  # Grupo empieza desde 1
                    row['orden_visita'],
                    row['regional'],
                    row['distrito'],
                    row['zona'],
                    row['sector'],
                    row['centro'],
                    row['latitud'],
                    row['longitud'],
                )
            )
        messagebox.showinfo("Éxito", "Tabla de grupos generada con éxito.")
    except Exception as e:
        messagebox.showerror("Error", f"Error al generar la tabla de grupos: {e}")


# ----------- INTERFAZ GRÁFICA -----------

# Botón para generar la tabla de grupos después de "Crear Grupos/Rutas"
btn_generar_tabla_grupos = tk.Button(
    frame_rutas, text="Generar Tabla de Grupos", command=generar_tabla_grupos
)
btn_generar_tabla_grupos.grid(row=2, column=0, columnspan=2, padx=5, pady=5)

# Frame para mostrar la tabla de grupos creados
frame_tabla_grupos = tk.Frame(root, bg="#ADD8E6", relief="solid", bd=2)
frame_tabla_grupos.grid(row=3, column=1, padx=10, pady=10, sticky="nsew")

# Configuración de la tabla
tree_scroll_y = tk.Scrollbar(frame_tabla_grupos, orient="vertical")
tree_scroll_y.pack(side="right", fill="y")

tree_scroll_x = tk.Scrollbar(frame_tabla_grupos, orient="horizontal")
tree_scroll_x.pack(side="bottom", fill="x")

'''
tree_grupos = ttk.Treeview(
    frame_tabla_grupos,
    columns=("Grupo", "Centro", "Regional", "Distrito", "Zona", "Sector", "Centro Nombre", "Latitud", "Longitud"),
    show="headings",
    yscrollcommand=tree_scroll_y.set,
    xscrollcommand=tree_scroll_x.set,
)
'''
tree_grupos = ttk.Treeview(
    frame_tabla_grupos,
    columns=("Grupo", "Centro", "Regional", "Distrito", "Zona", "Sector", "Centro Nombre"),
    show="headings",
    yscrollcommand=tree_scroll_y.set,
    xscrollcommand=tree_scroll_x.set,
)


# Configuración de encabezados
tree_grupos.heading("Grupo", text="Grupo")
tree_grupos.heading("Centro", text="Centro (# Orden)")
tree_grupos.heading("Regional", text="Regional")
tree_grupos.heading("Distrito", text="Distrito")
tree_grupos.heading("Zona", text="Zona")
tree_grupos.heading("Sector", text="Sector")
tree_grupos.heading("Centro Nombre", text="Centro")
#tree_grupos.heading("Latitud", text="Latitud")
#tree_grupos.heading("Longitud", text="Longitud")

# Ajustar ancho de columnas
tree_grupos.column("Grupo", anchor="center", width=80)
tree_grupos.column("Centro", anchor="center", width=120)
tree_grupos.column("Regional", anchor="center", width=150)
tree_grupos.column("Distrito", anchor="center", width=150)
tree_grupos.column("Zona", anchor="center", width=120)
tree_grupos.column("Sector", anchor="center", width=120)
tree_grupos.column("Centro Nombre", anchor="center", width=250)
#tree_grupos.column("Latitud", anchor="center", width=120)
#tree_grupos.column("Longitud", anchor="center", width=120)

tree_grupos.pack(fill="both", expand=True)

tree_scroll_y.config(command=tree_grupos.yview)
tree_scroll_x.config(command=tree_grupos.xview)

# Función para exportar la tabla de grupos a Excel
def exportar_tabla_grupos():
    global df_muestra
    if df_muestra.empty:
        messagebox.showwarning("Advertencia", "No hay grupos creados para exportar.")
        return

    try:
        # Ruta para guardar el archivo
        file_path = filedialog.asksaveasfilename(
            defaultextension=".xlsx", filetypes=[("Excel files", "*.xlsx")]
        )
        if file_path:
            # Seleccionar solo las columnas necesarias para exportar
            columnas_exportar = [
                'grupo', 'orden_visita', 'regional', 'distrito', 'zona', 'sector', 'centro', 'latitud', 'longitud'
            ]
            df_exportar = df_muestra[columnas_exportar].copy()
            df_exportar['grupo'] += 1  # Ajustar el grupo para que empiece desde 1
            df_exportar.rename(columns={
                'grupo': 'Grupo',
                'orden_visita': 'Centro (# Orden)',
                'regional': 'Regional',
                'distrito': 'Distrito',
                'zona': 'Zona',
                'sector': 'Sector',
                'centro': 'Centro',
                'latitud': 'Latitud',
                'longitud': 'Longitud'
            }, inplace=True)

            # Exportar a Excel
            df_exportar.to_excel(file_path, index=False)
            messagebox.showinfo("Éxito", "La tabla de grupos ha sido exportada con éxito.")
    except Exception as e:
        messagebox.showerror("Error", f"Error al exportar la tabla de grupos: {e}")


# ----------- INTERFAZ GRÁFICA -----------

# Botón para exportar la tabla de grupos a Excel
btn_exportar_tabla_grupos = tk.Button(
    frame_rutas, text="Exportar Tabla a Excel", command=exportar_tabla_grupos
)
btn_exportar_tabla_grupos.grid(row=3, column=0, columnspan=2, padx=5, pady=5)



# Footer
footer_frame = tk.Frame(root, bg="#005580", height=50)
footer_frame.grid(row=5, column=0, columnspan=2, sticky="ew")
footer_frame.grid_propagate(False)

footer_label = tk.Label(
    footer_frame,
    text="Copyright © Instituto Dominicano de Evaluación e Investigación de la Calidad Educativa IDEICE 2024",
    font=("Arial", 10),
    fg="white",
    bg="#005580",
)
footer_label.pack(pady=10)

# Configuración de la cuadrícula para que se ajuste el contenido
root.grid_rowconfigure(3, weight=1)
root.grid_columnconfigure(0, weight=1)

# Ejecutar el bucle principal de Tkinter
root.mainloop()
