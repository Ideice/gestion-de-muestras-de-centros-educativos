# Gestión de Muestras - Evaluación e Investigación Educativa

Este proyecto es una aplicación de escritorio desarrollada en Python con **Tkinter** para la gestión de muestras estadísticas en el ámbito de la evaluación e investigación educativa. La aplicación permite realizar cálculos de tamaño de muestra, selección de tipos de muestra (aleatoria simple, sistemática y estratificada), y la generación de rutas óptimas para la recolección de datos geolocalizados.

## Características Principales

1. **Carga de Datos**:
   - Permite cargar archivos Excel con información de centros educativos y datos asociados.
   - Campos obligatorios en el archivo Excel:
     - `latitud`: Coordenada geográfica de latitud.
     - `longitud`: Coordenada geográfica de longitud.
     - `regional`: Nombre o código de la regional educativa.
     - `centro`: Nombre del centro educativo.
     - `sector`: Clasificación del sector (público/privado).
     - `zona`: Zona geográfica del centro.
     - `nivel_educativo`: Nivel educativo que abarca el centro.

2. **Cálculo de Tamaño de Muestra**:
   - Calcula el tamaño de muestra según nivel de confianza y margen de error.

3. **Tipos de Muestra**:
   - Aleatoria Simple.
   - Aleatoria Simple Sistemática.
   - Estratificada (Regional, Zona, Sector, Nivel Educativo).
   - Estratificada Sistemática.

4. **Generación de Rutas y Grupos**:
   - Divide los datos en grupos usando el algoritmo de clustering K-Means.
   - Calcula la ruta óptima dentro de cada grupo basado en la distancia más corta (heurística de vecino más cercano).
   - Visualiza las rutas en mapas interactivos generados con **Folium**.

5. **Exportación**:
   - Exporta las muestras y grupos generados a archivos Excel.
   - Guarda mapas de centros educativos y rutas en formato HTML.

6. **Interfaz Gráfica**:
   - Construida con Tkinter, incluye elementos interactivos como botones, tablas, y menús desplegables.

## Tecnologías Utilizadas

- **Python**:
  - Tkinter para la interfaz gráfica.
  - Pandas para el manejo de datos.
  - Scipy para cálculos estadísticos.
  - Folium para mapas interactivos.
  - Scikit-learn para clustering (K-Means).

- **Dependencias**:
  - pandas
  - scipy
  - numpy
  - sklearn
  - folium

## Requisitos

1. Python 3.8 o superior.
2. Las siguientes bibliotecas instaladas:
   ```bash
   pip install pandas scipy numpy scikit-learn folium
