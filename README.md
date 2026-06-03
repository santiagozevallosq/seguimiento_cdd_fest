# Dashboard de Control de Avances - Proyectos CDD - FEST

Este proyecto es un panel de control interactivo diseñado para monitorear el avance de ejecución física y financiera de los proyectos CDD - FEST, facilitando la identificación de rezagos, la comparación de progresos y la detección de proyectos que requieren reprogramación.

El dashboard está desarrollado utilizando únicamente **tecnologías web estándar (HTML5, CSS3, JavaScript puro)** y librerías externas ligeras a través de CDNs, lo que permite un rendimiento óptimo y un despliegue sin servidor inmediato en plataformas como Vercel.

---

## 🚀 Características Principales

- **Lectura directa de Excel:** Procesa el archivo excel mediante la librería [SheetJS (XLSX)](https://sheetjs.com/).
- **Visualización dinámica:** Gráficos del ranking del avance físico, certificado y devengado por CITE, además de la proporción de reprogramaciones, usando [Chart.js](https://www.chartjs.org/).
- **KPIs Automáticos:** Tarjetas de rendimiento con el número total de proyectos y promedios actualizados en tiempo real según filtros.
- **Filtros en cascada interactivos:** Filtros por código de proyecto, CITE, PET y reprogramación, que se actualizan de forma inteligente y combinada como en sistemas BI (Power BI o Looker Studio).
- **Tabla de Detalle con Formato Condicional:** 
  - Mapa de calor (*heatmap*) en celdas de avance físico y financiero.
  - Columna de **Alerta** con lógica operativa-financiera calculada al instante.

---

## 🛠️ Estructura del Proyecto

```text
seguimiento_cdd_fest/
├── index.html          # Interfaz de usuario (Estructura y vinculación de CDNs)
├── styles.css          # Estilos ejecutivos premium, colores institucionales y responsive
├── app.js              # Lógica de procesamiento de datos, filtros y renderizado de gráficos/tablas
├── data.xlsx           # Copia local del archivo Excel para carga automática
├── package.json        # Configuración para levantar el servidor local rápidamente
└── README.md           # Guía de uso y despliegue (Este archivo)
```

---

## 💻 Ejecución Local

Dado que el navegador restringe la lectura de archivos locales (`fetch` a un archivo local como `data.xlsx`) por políticas de seguridad (CORS), es necesario ejecutar el dashboard a través de un servidor web local simple.

### Pasos para iniciar:

1. **Iniciar el servidor local con Python:**
   Dado que Python está instalado en tu sistema, abre una terminal (PowerShell o CMD) en la carpeta raíz del proyecto y ejecuta:
   ```bash
   python -m http.server 8080
   ```
   *(También puedes usar `npm start` si instalas Node.js más adelante).*

2. **Ver en el navegador:**
   Abre tu navegador de preferencia e ingresa a la dirección:
   [http://localhost:8080](http://localhost:8080)

---

## 📊 Configuración de Carga del Archivo Excel

El código JavaScript en `app.js` ofrece dos alternativas integradas para alimentar el dashboard:

### Alternativa A: Carga Manual (Subir Archivo)
Permite al usuario subir cualquier archivo Excel compatible directamente desde la interfaz web del dashboard pulsando el botón **"Subir Excel"** en el encabezado.
*Para usar únicamente esta alternativa:* Comenta la línea `cargarExcelAutomatico("data.xlsx");` en el evento `DOMContentLoaded` de `app.js` y remueve la clase `hidden` del elemento `#no-data-state` en `index.html` para mostrar el aviso inicial de carga.

### Alternativa B: Carga Automática (Predeterminada)
Lee de manera automática un archivo llamado `data.xlsx` ubicado en la raíz del proyecto.
*Para actualizar los datos:* Solo debes reemplazar el archivo `data.xlsx` de la raíz por tu nuevo archivo de corte con el mismo nombre y recargar la página.

---

## 🎨 Lógica de Alertas en Tabla de Detalle

Las alertas se asignan a cada proyecto siguiendo esta prioridad técnica:
1. **Crítico operativo (Fondo Rojo):** Avance físico $< 25\%$ y Devengado financiero $< 25\%$.
2. **Riesgo financiero (Fondo Naranja):** Certificado financiero $> 50\%$ y Devengado financiero $< 25\%$.
3. **Desfase físico-financiero (Fondo Amarillo):** Certificado financiero $-$ Avance físico $> 30\%$.
4. **Sin alerta (Fondo Verde):** El proyecto cumple satisfactoriamente con los parámetros.

---

## ☁️ Despliegue en Vercel

Desplegar este dashboard en Vercel es sumamente rápido puesto que es un sitio web estático:

### Método 1: Vercel CLI (Línea de comandos)
1. Instala Vercel de forma global si no lo tienes:
   ```bash
   npm install -g vercel
   ```
2. Ejecuta el comando `vercel` en la carpeta raíz del proyecto:
   ```bash
   vercel
   ```
3. Sigue las instrucciones interactivas en la consola. Vercel detectará el proyecto y te entregará una URL pública de inmediato.

### Método 2: Integración con GitHub (Recomendado para actualización continua)
1. Sube este proyecto a un repositorio en tu cuenta de GitHub (ej. `mi-usuario/seguimiento-cdd-fest`).
2. Entra a tu cuenta en [Vercel](https://vercel.com/) y haz clic en **"Add New"** > **"Project"**.
3. Importa tu repositorio recién creado.
4. Haz clic en **"Deploy"**. Vercel configurará y publicará el dashboard. Cada cambio que subas a tu repositorio de GitHub se reflejará automáticamente en tu sitio desplegado.
