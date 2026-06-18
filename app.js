// Variables globales para almacenar los datos de la aplicación y las instancias de gráficos
let allData = [];
let filteredData = [];
let charts = {
    physical: null,
    committed: null,
    accrued: null
};

// Registrar el plugin de etiquetas de datos para todos los gráficos
Chart.register(ChartDataLabels);

// Configuración de la carga inicial del archivo Excel y control de acceso
document.addEventListener("DOMContentLoaded", () => {
    inicializarLogin();
    inicializarEventos();
    
    // ALTERNATIVA B: Carga automática de "data.xlsx" en la raíz del proyecto
    // Para activar la ALTERNATIVA A (solo carga manual), comente o borre el bloque de código abajo
    // y elimine la clase 'hidden' del contenedor de estado #no-data-state en index.html
    cargarExcelAutomatico("data.xlsx");
});

// Función para gestionar la autenticación local (Opción A)
function inicializarLogin() {
    const overlay = document.getElementById("login-overlay");
    const form = document.getElementById("login-form");
    const errorEl = document.getElementById("login-error");
    const usernameInput = document.getElementById("login-username");
    const passwordInput = document.getElementById("login-password");

    // Verificar si ya está autenticado en la sesión actual
    if (sessionStorage.getItem("authenticated") === "true") {
        overlay.style.display = "none";
        return;
    }

    form.addEventListener("submit", (e) => {
        e.preventDefault();
        
        const username = usernameInput.value.trim();
        const password = passwordInput.value;

        // Credenciales: dese / dese
        if (username === "dese" && password === "dese") {
            sessionStorage.setItem("authenticated", "true");
            errorEl.classList.add("hidden");
            
            // Efecto de fade-out
            overlay.classList.add("fade-out");
            setTimeout(() => {
                overlay.style.display = "none";
            }, 400);
        } else {
            errorEl.classList.remove("hidden");
            passwordInput.value = "";
            passwordInput.focus();
        }
    });
}

// Función para inicializar los event listeners de los filtros y el input de archivo
function inicializarEventos() {
    // Input de archivo para Alternativa A
    const fileInput = document.getElementById("excel-file-input");
    fileInput.addEventListener("change", cargarExcelManual);

    // Eventos para filtros superiores
    document.getElementById("filter-cite").addEventListener("change", (e) => {
        document.getElementById("table-filter-cite").value = e.target.value;
        aplicarFiltros();
    });
    document.getElementById("filter-pet").addEventListener("change", (e) => {
        document.getElementById("table-filter-pet").value = e.target.value;
        aplicarFiltros();
    });
    document.getElementById("filter-reprog").addEventListener("change", (e) => {
        document.getElementById("table-filter-reprog").value = e.target.value;
        aplicarFiltros();
    });

    // Eventos para filtros de cabecera de la tabla
    document.getElementById("table-filter-cite").addEventListener("change", (e) => {
        document.getElementById("filter-cite").value = e.target.value;
        aplicarFiltros();
    });
    document.getElementById("table-filter-pet").addEventListener("change", (e) => {
        document.getElementById("filter-pet").value = e.target.value;
        aplicarFiltros();
    });
    document.getElementById("table-filter-cod-proy").addEventListener("change", () => {
        aplicarFiltros();
    });
    document.getElementById("table-filter-reprog").addEventListener("change", (e) => {
        document.getElementById("filter-reprog").value = e.target.value;
        aplicarFiltros();
    });

    // Botón de limpiar filtros
    document.getElementById("btn-clear-filters").addEventListener("click", limpiarFiltros);
}

// ALTERNATIVA B: Lectura automática del archivo Excel desde una ruta local
function cargarExcelAutomatico(url) {
    mostrarCarga(true);
    fetch(url)
        .then(response => {
            if (!response.ok) {
                throw new Error("No se pudo cargar el archivo Excel automáticamente.");
            }
            return response.arrayBuffer();
        })
        .then(data => {
            procesarExcel(data, "data.xlsx");
        })
        .catch(error => {
            console.warn("Carga automática fallida. Esperando carga manual del usuario.", error);
            mostrarCarga(false);
            mostrarEstadoSinDatos(true);
        });
}

// ALTERNATIVA A: Carga de archivo manual desde el input tipo file
function cargarExcelManual(event) {
    const file = event.target.files[0];
    if (!file) return;

    mostrarCarga(true);
    mostrarEstadoSinDatos(false);
    
    const reader = new FileReader();
    reader.onload = function(e) {
        const data = new Uint8Array(e.target.result);
        procesarExcel(data, file.name);
    };
    reader.readAsArrayBuffer(file);
}

// Función principal de procesamiento de datos con SheetJS
function procesarExcel(arrayBuffer, fileName) {
    try {
        const workbook = XLSX.read(arrayBuffer, { type: "array" });
        
        // Buscar la hoja llamada "Ejecutivo" o usar la primera hoja disponible
        let sheetName = "Ejecutivo";
        if (!workbook.SheetNames.includes(sheetName)) {
            sheetName = workbook.SheetNames[0];
            console.warn(`Hoja 'Ejecutivo' no encontrada. Leyendo la hoja: ${sheetName}`);
        }

        const worksheet = workbook.Sheets[sheetName];
        const rawJson = XLSX.utils.sheet_to_json(worksheet, { defval: null });

        if (!rawJson || rawJson.length === 0) {
            throw new Error("El archivo Excel no tiene datos válidos.");
        }

        // Limpiar y transformar los datos
        allData = limpiarDatos(rawJson);
        filteredData = [...allData];

        // Actualizar etiqueta del archivo en el header
        document.getElementById("file-name-label").textContent = fileName;

        // Inicializar filtros y renderizar vistas
        inicializarOpcionesFiltros();
        actualizarDashboard();
        
        mostrarCarga(false);
    } catch (error) {
        alert("Error al procesar el archivo Excel: " + error.message);
        console.error(error);
        mostrarCarga(false);
        if (allData.length === 0) {
            mostrarEstadoSinDatos(true);
        }
    }
}

// Limpieza y transformación de datos según requisitos
function limpiarDatos(rawJson) {
    return rawJson
        .filter(row => {
            // Filtrar filas vacías o sin identificador de número/código
            return row["n"] !== null && (row["codigo_proyecto"] || row["codigo_proyec"]);
        })
        .map(row => {
            // 1. Mapear códigos de proyecto
            const codigo = (row["codigo_proyecto"] || row["codigo_proyec"] || "").toString().trim();
            
            // 2. CITE y PET
            const cite = (row["cite"] || "Sin CITE").toString().trim();
            const pet = (row["pet"] || "Sin PET").toString().trim();

            // 3. Limpieza de montos financieros
            const financ_prog = parseCleanNumber(row["financ_prog"]);
            const financ_cert = parseCleanNumber(row["financ_cert"]);
            const financ_compr = parseCleanNumber(row["financ_compr"]);
            const financ_deven = parseCleanNumber(row["financ_deven"]);

            // 4. Limpieza y transformación de porcentajes
            const fisica_prog = parseCleanNumber(row["fisica_prog"]);
            const fisica_ejec = parseCleanNumber(row["fisica_ejec"]);
            
            const fisica_pct = parsePercentage(row["fisica_pct"]);
            const financ_cert_pct = parsePercentage(row["financ_cert_pct"]);
            const financ_compr_pct = parsePercentage(row["financ_compr_pct"]);
            const financ_deven_pct = parsePercentage(row["financ_deven_pct"]);

            // 5. Normalizar "req_reprogr"
            let req_reprogr = (row["req_reprogr"] || "No").toString().trim();
            if (/^s[ií]/i.test(req_reprogr)) {
                req_reprogr = "Sí";
            } else {
                req_reprogr = "No";
            }

            // 5.5 Nuevas variables de servicios y unidades productivas (celda I y J)
            const servicio = parseCleanNumber(row["servicio"]);
            const up = parseCleanNumber(row["up"]);

            // 6. Nuevas variables cualitativas
            let mma_estado = (row["mma_estado"] || "").toString().trim();
            if (mma_estado.toLowerCase() === "sin especificar") {
                mma_estado = "";
            }
            let detalles = (row["Incidencias reportadas"] || row["incidencias reportadas"] || row["Detalles"] || row["detalles"] || row["detalle"] || "").toString().trim();
            if (detalles.toLowerCase() === "sin especificar") {
                detalles = "";
            }

            return {
                n: row["n"],
                codigo_proyecto: codigo,
                cite: cite,
                pet: pet,
                fisica_prog: fisica_prog,
                fisica_ejec: fisica_ejec,
                fisica_pct: fisica_pct,
                financ_prog: financ_prog,
                financ_cert: financ_cert,
                financ_compr: financ_compr,
                financ_deven: financ_deven,
                financ_cert_pct: financ_cert_pct,
                financ_compr_pct: financ_compr_pct,
                financ_deven_pct: financ_deven_pct,
                req_reprogr: req_reprogr,
                mma_estado: mma_estado,
                detalles: detalles,
                servicio: servicio,
                up: up
            };
        });
}

// Convertir monedas o cadenas numéricas a float
function parseCleanNumber(val) {
    if (val === undefined || val === null) return 0;
    if (typeof val === "number") return val;
    
    // Limpieza de símbolos monetarios comunes y comas de formato
    let cleaned = val.toString()
        .replace(/S\/\.?/g, "")
        .replace(/,/g, "")
        .trim();
    
    let num = parseFloat(cleaned);
    return isNaN(num) ? 0 : num;
}

// Convertir porcentajes en cadena o decimales a escala 0-100
function parsePercentage(val) {
    if (val === undefined || val === null) return 0;
    if (typeof val === "number") {
        // Si el valor viene como decimal en el rango [0, 1] ej: 0.4285
        // Se asume que representa un porcentaje decimal y se multiplica por 100.
        // Si es mayor a 1, ya viene como 42.85 y se mantiene.
        if (val > 0 && val <= 1) {
            return val * 100;
        }
        return val;
    }

    // Si viene como string tipo "42.9%" o "0.429"
    let cleaned = val.toString().replace(/%/g, "").trim();
    let num = parseFloat(cleaned);
    if (isNaN(num)) return 0;

    // Tratar números decimales en formato de texto
    if (num > 0 && num <= 1) {
        return num * 100;
    }
    return num;
}

// Formatear montos en Soles peruanos
function formatearSoles(value) {
    return new Intl.NumberFormat("es-PE", {
        style: "currency",
        currency: "PEN",
        minimumFractionDigits: 0,
        maximumFractionDigits: 0
    }).format(value);
}

// Cargar las opciones iniciales en los filtros dropdown
function inicializarOpcionesFiltros() {
    actualizarFiltrosCascada();
}

// Actualizar los dropdowns basados en la selección actual para crear el efecto en cascada
function actualizarFiltrosCascada() {
    const selCite = document.getElementById("filter-cite");
    const selPet = document.getElementById("filter-pet");
    const selReprog = document.getElementById("filter-reprog");
    const selTableCite = document.getElementById("table-filter-cite");
    const selTablePet = document.getElementById("table-filter-pet");
    const selTableCod = document.getElementById("table-filter-cod-proy");
    const selTableReprog = document.getElementById("table-filter-reprog");

    // Guardar los valores previamente seleccionados
    const valCite = selCite.value;
    const valPet = selPet.value;
    const valCod = selTableCod.value;
    const valReprog = selReprog.value;

    // Generar opciones filtrando los datos del resto de componentes
    
    // CITEs disponibles considerando los otros filtros
    const dataForCite = allData.filter(d => 
        (!valPet || d.pet === valPet) &&
        (!valCod || d.codigo_proyecto === valCod) &&
        (!valReprog || d.req_reprogr === valReprog)
    );
    const optCite = [...new Set(dataForCite.map(d => d.cite))].sort();

    // PETs disponibles considerando los otros filtros
    const dataForPet = allData.filter(d => 
        (!valCite || d.cite === valCite) &&
        (!valCod || d.codigo_proyecto === valCod) &&
        (!valReprog || d.req_reprogr === valReprog)
    );
    const optPet = [...new Set(dataForPet.map(d => d.pet))].sort();

    // Códigos de proyecto disponibles considerando los otros filtros
    const dataForCod = allData.filter(d => 
        (!valCite || d.cite === valCite) &&
        (!valPet || d.pet === valPet) &&
        (!valReprog || d.req_reprogr === valReprog)
    );
    const optCod = [...new Set(dataForCod.map(d => d.codigo_proyecto))].sort();

    // Reprogramaciones disponibles considerando los otros filtros
    const dataForReprog = allData.filter(d => 
        (!valCite || d.cite === valCite) &&
        (!valPet || d.pet === valPet) &&
        (!valCod || d.codigo_proyecto === valCod)
    );
    const optReprog = [...new Set(dataForReprog.map(d => d.req_reprogr))].sort();

    // Volver a rellenar los selects preservando los valores
    rellenarSelect(selCite, optCite, valCite, "Todos los CITEs");
    rellenarSelect(selPet, optPet, valPet, "Todos los PET");
    rellenarSelect(selReprog, optReprog, valReprog, "Todos");
    
    rellenarSelect(selTableCite, optCite, valCite, "Todos");
    rellenarSelect(selTablePet, optPet, valPet, "Todos");
    rellenarSelect(selTableCod, optCod, valCod, "Todos");
    rellenarSelect(selTableReprog, optReprog, valReprog, "Todos");
}

function rellenarSelect(selectElement, options, selectedValue, defaultText) {
    selectElement.innerHTML = `<option value="">${defaultText}</option>`;
    options.forEach(opt => {
        if (opt !== undefined && opt !== null && opt !== "") {
            const selected = opt.toString() === selectedValue ? "selected" : "";
            const optionElem = document.createElement("option");
            optionElem.value = opt;
            optionElem.textContent = opt;
            if (selected) optionElem.selected = true;
            selectElement.appendChild(optionElem);
        }
    });
}

// Aplicar filtros seleccionados y actualizar vistas
function aplicarFiltros() {
    const valCite = document.getElementById("filter-cite").value;
    const valPet = document.getElementById("filter-pet").value;
    const valCod = document.getElementById("table-filter-cod-proy").value;
    const valReprog = document.getElementById("filter-reprog").value;

    // Filtrar el conjunto de datos completo
    filteredData = allData.filter(d => {
        return (!valCite || d.cite === valCite) &&
               (!valPet || d.pet === valPet) &&
               (!valCod || d.codigo_proyecto === valCod) &&
               (!valReprog || d.req_reprogr === valReprog);
    });

    // Actualizar el estado de los filtros en cascada
    actualizarFiltrosCascada();
    
    // Refrescar KPIs, Gráficos y Tabla
    actualizarDashboard();
}

// Limpiar filtros restableciendo todos los valores y datos
function limpiarFiltros() {
    document.getElementById("filter-cite").value = "";
    document.getElementById("filter-pet").value = "";
    document.getElementById("filter-reprog").value = "";
    document.getElementById("table-filter-cite").value = "";
    document.getElementById("table-filter-pet").value = "";
    document.getElementById("table-filter-cod-proy").value = "";
    document.getElementById("table-filter-reprog").value = "";

    filteredData = [...allData];
    
    actualizarFiltrosCascada();
    actualizarDashboard();
}

// Renderizar todos los elementos visuales del Dashboard
function actualizarDashboard() {
    actualizarKPIs();
    actualizarGraficos();
    actualizarTabla();
}

// Calcular y renderizar KPIs
function actualizarKPIs() {
    const totalProy = filteredData.length;
    let totalProg = 0;
    let totalCert = 0;
    let totalCompr = 0;
    let totalDeven = 0;
    let totalReprog = 0;
    let pctCert = 0;
    let pctCompr = 0;
    let pctDeven = 0;

    let totalFisicaProg = 0;
    let totalFisicaEjec = 0;
    let pctFisicaEjec = 0;
    let totalServicio = 0;
    let totalUp = 0;

    if (totalProy > 0) {
        // Sumas de montos
        totalProg = filteredData.reduce((acc, curr) => acc + curr.financ_prog, 0);
        totalCert = filteredData.reduce((acc, curr) => acc + curr.financ_cert, 0);
        totalCompr = filteredData.reduce((acc, curr) => acc + curr.financ_compr, 0);
        totalDeven = filteredData.reduce((acc, curr) => acc + curr.financ_deven, 0);

        // Reprogramaciones
        totalReprog = filteredData.filter(d => d.req_reprogr === "Sí").length;

        // Porcentajes globales
        pctCert = totalProg > 0 ? (totalCert / totalProg) * 100 : 0;
        pctCompr = totalProg > 0 ? (totalCompr / totalProg) * 100 : 0;
        pctDeven = totalProg > 0 ? (totalDeven / totalProg) * 100 : 0;

        // Nuevos KPI
        totalFisicaProg = filteredData.reduce((acc, curr) => acc + curr.fisica_prog, 0);
        totalFisicaEjec = filteredData.reduce((acc, curr) => acc + curr.fisica_ejec, 0);
        pctFisicaEjec = totalFisicaProg > 0 ? (totalFisicaEjec / totalFisicaProg) * 100 : 0;
        totalServicio = filteredData.reduce((acc, curr) => acc + (curr.servicio || 0), 0);
        totalUp = filteredData.reduce((acc, curr) => acc + (curr.up || 0), 0);
    }

    // Renderizar nuevos KPIs
    document.getElementById("kpi-val-projects").textContent = totalProy;
    document.getElementById("kpi-val-prog").textContent = formatearSoles(totalProg);
    document.getElementById("kpi-val-cert").textContent = formatearSoles(totalCert);
    document.getElementById("kpi-val-cert-pct").textContent = `${pctCert.toFixed(1)}%`;
    document.getElementById("kpi-val-compr").textContent = formatearSoles(totalCompr);
    document.getElementById("kpi-val-compr-pct").textContent = `${pctCompr.toFixed(1)}%`;
    document.getElementById("kpi-val-deven").textContent = formatearSoles(totalDeven);
    document.getElementById("kpi-val-deven-pct").textContent = `${pctDeven.toFixed(1)}%`;
    document.getElementById("kpi-val-reprog").textContent = totalReprog;

    // Renderizar nuevos KPIs
    document.getElementById("kpi-val-fisica-prog").textContent = totalFisicaProg;
    document.getElementById("kpi-val-fisica-ejec").textContent = totalFisicaEjec;
    document.getElementById("kpi-val-fisica-ejec-pct").textContent = `(${pctFisicaEjec.toFixed(1)}%)`;
    document.getElementById("kpi-val-servicio").textContent = totalServicio;
    document.getElementById("kpi-val-up").textContent = totalUp;
}

// Funciones auxiliares para semaforización de colores (Verde >= 80%, Ambar >= 50%, Rojo < 50%)
function obtenerColorSemaforo(val) {
    if (val >= 80) {
        return {
            bg: "rgba(34, 197, 94, 0.85)",     // Verde claro
            border: "rgba(34, 197, 94, 1)"
        };
    } else if (val >= 50) {
        return {
            bg: "rgba(245, 158, 11, 0.85)",     // Ámbar
            border: "rgba(245, 158, 11, 1)"
        };
    } else {
        return {
            bg: "rgba(239, 68, 68, 0.85)",     // Rojo claro
            border: "rgba(239, 68, 68, 1)"
        };
    }
}

function obtenerEstiloCeldaSemaforo(val) {
    if (val >= 80) {
        return "background-color: var(--alert-ok-bg); color: var(--alert-ok-text); font-weight: 600;";
    } else if (val >= 50) {
        return "background-color: var(--alert-riesgo-bg); color: var(--alert-riesgo-text); font-weight: 600;";
    } else {
        return "background-color: var(--alert-critico-bg); color: var(--alert-critico-text); font-weight: 600;";
    }
}

// Renderizar gráficos de Chart.js
function actualizarGraficos() {
    destruirGraficosExistentes();

    if (filteredData.length === 0) return;

    // Preparar datos ordenados para gráficos horizontales por CITE/Código (usando CITE)
    // Agrupamos por CITE para mejor visualización de avance si hay duplicados
    const citeData = agruparPorCite(filteredData);

    // Ajustar altura de los contenedores de los gráficos de barras horizontales según la cantidad de CITEs
    const dynamicHeight = Math.max(240, citeData.length * 32 + 35);
    document.getElementById("chart-physical").parentElement.parentElement.style.height = `${dynamicHeight}px`;
    document.getElementById("chart-committed").parentElement.parentElement.style.height = `${dynamicHeight}px`;
    document.getElementById("chart-accrued").parentElement.parentElement.style.height = `${dynamicHeight}px`;

    // Gráfico 1: % Ejecución Física (Barras Horizontales - Ordenado de mayor a menor)
    const sortedPhysical = [...citeData].sort((a, b) => b.fisica_pct - a.fisica_pct);
    const labelsPhysical = sortedPhysical.map(d => d.cite);
    const dataPhysical = sortedPhysical.map(d => d.fisica_pct);
    
    const colorsPhysical = dataPhysical.map(val => obtenerColorSemaforo(val).bg);
    const bordersPhysical = dataPhysical.map(val => obtenerColorSemaforo(val).border);

    const ctxPhysical = document.getElementById("chart-physical").getContext("2d");
    charts.physical = new Chart(ctxPhysical, {
        type: "bar",
        data: {
            labels: labelsPhysical,
            datasets: [{
                label: "% Avance Físico",
                data: dataPhysical,
                backgroundColor: colorsPhysical,
                borderColor: bordersPhysical,
                borderWidth: 1.5,
                borderRadius: 4
            }]
        },
        options: obtenerConfiguracionGraficoBarras(true, "physical") // true para horizontal
    });

    // Gráfico 2: % Comprometido (Barras Horizontales - Ordenado de mayor a menor)
    const sortedCommitted = [...citeData].sort((a, b) => b.financ_compr_pct - a.financ_compr_pct);
    const labelsCommitted = sortedCommitted.map(d => d.cite);
    const dataCommitted = sortedCommitted.map(d => d.financ_compr_pct);

    const colorsCommitted = dataCommitted.map(val => obtenerColorSemaforo(val).bg);
    const bordersCommitted = dataCommitted.map(val => obtenerColorSemaforo(val).border);

    const ctxCommitted = document.getElementById("chart-committed").getContext("2d");
    charts.committed = new Chart(ctxCommitted, {
        type: "bar",
        data: {
            labels: labelsCommitted,
            datasets: [{
                label: "% Comprometido Financiero",
                data: dataCommitted,
                monetaryData: sortedCommitted.map(d => d.financ_compr),
                backgroundColor: colorsCommitted,
                borderColor: bordersCommitted,
                borderWidth: 1.5,
                borderRadius: 4
            }]
        },
        options: obtenerConfiguracionGraficoBarras(true, "committed")
    });

    // Gráfico 3: % Devengado (Barras Horizontales - Ordenado de mayor a menor)
    const sortedAccrued = [...citeData].sort((a, b) => b.financ_deven_pct - a.financ_deven_pct);
    const labelsAccrued = sortedAccrued.map(d => d.cite);
    const dataAccrued = sortedAccrued.map(d => d.financ_deven_pct);

    const colorsAccrued = dataAccrued.map(val => obtenerColorSemaforo(val).bg);
    const bordersAccrued = dataAccrued.map(val => obtenerColorSemaforo(val).border);

    const ctxAccrued = document.getElementById("chart-accrued").getContext("2d");
    charts.accrued = new Chart(ctxAccrued, {
        type: "bar",
        data: {
            labels: labelsAccrued,
            datasets: [{
                label: "% Devengado Financiero",
                data: dataAccrued,
                monetaryData: sortedAccrued.map(d => d.financ_deven),
                backgroundColor: colorsAccrued,
                borderColor: bordersAccrued,
                borderWidth: 1.5,
                borderRadius: 4
            }]
        },
        options: obtenerConfiguracionGraficoBarras(true, "accrued")
    });
}

// Agrupar filas del Excel por CITE para los gráficos agregados
function agruparPorCite(data) {
    const groups = {};
    data.forEach(item => {
        if (!groups[item.cite]) {
            groups[item.cite] = {
                cite: item.cite,
                fisica_pct_sum: 0,
                financ_compr_pct_sum: 0,
                financ_deven_pct_sum: 0,
                financ_compr_sum: 0,
                financ_deven_sum: 0,
                count: 0
            };
        }
        groups[item.cite].fisica_pct_sum += item.fisica_pct;
        groups[item.cite].financ_compr_pct_sum += item.financ_compr_pct;
        groups[item.cite].financ_deven_pct_sum += item.financ_deven_pct;
        groups[item.cite].financ_compr_sum += item.financ_compr || 0;
        groups[item.cite].financ_deven_sum += item.financ_deven || 0;
        groups[item.cite].count++;
    });

    return Object.values(groups).map(g => ({
        cite: g.cite,
        fisica_pct: g.fisica_pct_sum / g.count,
        financ_compr_pct: g.financ_compr_pct_sum / g.count,
        financ_deven_pct: g.financ_deven_pct_sum / g.count,
        financ_compr: g.financ_compr_sum,
        financ_deven: g.financ_deven_sum
    }));
}

// Configuración común de ejes y leyendas para los gráficos de barras
function obtenerConfiguracionGraficoBarras(isHorizontal, chartType = "physical") {
    return {
        indexAxis: isHorizontal ? "y" : "x",
        responsive: true,
        maintainAspectRatio: false,
        layout: {
            padding: {
                left: 20, // Agregado padding izquierdo para evitar que se corten los nombres de los CITEs
                right: chartType === "physical" ? 40 : 110 // Más espacio a la derecha para que no se corten etiquetas con montos
            }
        },
        plugins: {
            legend: {
                display: false
            },
            tooltip: {
                callbacks: {
                    label: function(context) {
                        const pct = context.raw.toFixed(1);
                        if (chartType !== "physical") {
                            const dataset = context.dataset;
                            const monVal = dataset.monetaryData ? dataset.monetaryData[context.dataIndex] : 0;
                            return ` Avance: ${pct}% (${formatearSoles(monVal)})`;
                        }
                        return ` Avance: ${pct}%`;
                    }
                }
            },
            datalabels: {
                anchor: "end",
                align: "right",
                formatter: function(value, context) {
                    const pct = value.toFixed(1) + "%";
                    if (chartType !== "physical") {
                        const dataset = context.dataset;
                        const monVal = dataset.monetaryData ? dataset.monetaryData[context.dataIndex] : 0;
                        return `${pct} (${formatearSoles(monVal)})`;
                    }
                    return pct;
                },
                font: {
                    family: "Inter",
                    size: 9,
                    weight: "bold"
                },
                color: "#1e293b"
            }
        },
        scales: {
            x: {
                min: 0,
                max: 100,
                grid: {
                    color: "#f1f5f9"
                },
                ticks: {
                    font: { family: "Inter", size: 10 },
                    color: "#64748b",
                    callback: function(value) { return value + "%"; }
                }
            },
            y: {
                grid: {
                    display: false
                },
                ticks: {
                    font: { family: "Inter", size: 8.5 },
                    color: "#475569",
                    autoSkip: false
                }
            }
        }
    };
}

// Destruir gráficos anteriores para evitar errores y comportamientos erráticos
function destruirGraficosExistentes() {
    Object.keys(charts).forEach(key => {
        if (charts[key]) {
            charts[key].destroy();
            charts[key] = null;
        }
    });
}

// Renderizar filas de la Tabla de Detalle y aplicar Heatmap & Alertas
function actualizarTabla() {
    const tableBody = document.getElementById("table-body");
    const rowCountEl = document.getElementById("table-row-count");
    tableBody.innerHTML = "";
    
    rowCountEl.textContent = `${filteredData.length} items`;

    filteredData.forEach(row => {
        const tr = document.createElement("tr");
        
        // Crear celdas con estilos dinámicos tipo semáforo
        const styleFisico = obtenerEstiloCeldaSemaforo(row.fisica_pct);
        const styleCertificado = obtenerEstiloCeldaSemaforo(row.financ_cert_pct);
        const styleComprometido = obtenerEstiloCeldaSemaforo(row.financ_compr_pct);
        const styleDevengado = obtenerEstiloCeldaSemaforo(row.financ_deven_pct);

        tr.innerHTML = `
            <td class="text-center">${row.n}</td>
            <td style="font-weight: 600; color: var(--color-primary); white-space: nowrap;">${row.codigo_proyecto}</td>
            <td>${row.cite}</td>
            <td>${row.pet}</td>
            <td class="text-right">${row.fisica_prog}</td>
            <td class="text-right">${row.fisica_ejec}</td>
            <td class="text-right" style="${styleFisico}">${row.fisica_pct.toFixed(1)}%</td>
            <td class="text-right" style="white-space: nowrap;">${formatearSoles(row.financ_prog)}</td>
            <td class="text-right" style="white-space: nowrap;">${formatearSoles(row.financ_cert)}</td>
            <td class="text-right" style="white-space: nowrap;">${formatearSoles(row.financ_compr)}</td>
            <td class="text-right" style="white-space: nowrap;">${formatearSoles(row.financ_deven)}</td>
            <td class="text-right" style="${styleCertificado}">${row.financ_cert_pct.toFixed(1)}%</td>
            <td class="text-right" style="${styleComprometido}">${row.financ_compr_pct.toFixed(1)}%</td>
            <td class="text-right" style="${styleDevengado}">${row.financ_deven_pct.toFixed(1)}%</td>
            <td class="text-center">
                <span class="badge ${row.req_reprogr === "Sí" ? "btn-primary" : "badge-light"}" style="font-size: 0.7rem; padding: 0.15rem 0.45rem;">
                    ${row.req_reprogr}
                </span>
            </td>
            <td>${row.mma_estado}</td>
            <td>${row.detalles}</td>
        `;
        
        tableBody.appendChild(tr);
    });
}

// Calcular Alertas de acuerdo con la lógica operativa-financiera
function calcularAlerta(row) {
    const fisica = row.fisica_pct;
    const cert = row.financ_compr_pct;
    const deven = row.financ_deven_pct;

    // Lógica 1: Crítico operativo (fisica_pct < 25 y financ_deven_pct < 25)
    if (fisica < 25 && deven < 25) {
        return {
            texto: "Crítico operativo",
            clase: "alert-critico"
        };
    }

    // Lógica 2: Riesgo financiero (financ_compr_pct > 50 y financ_deven_pct < 25)
    if (cert > 50 && deven < 25) {
        return {
            texto: "Riesgo financiero",
            clase: "alert-riesgo"
        };
    }

    // Lógica 3: Desfase físico-financiero (financ_compr_pct - fisica_pct > 30)
    if ((cert - fisica) > 30) {
        return {
            texto: "Desfase físico-financiero",
            clase: "alert-desfase"
        };
    }

    // Por defecto: Sin alerta
    return {
        texto: "Sin alerta",
        clase: "alert-ok"
    };
}

// Mostrar/Ocultar indicador de carga de datos
function mostrarCarga(mostrar) {
    const loader = document.getElementById("loading-state");
    if (mostrar) {
        loader.classList.remove("hidden");
    } else {
        loader.classList.add("hidden");
    }
}

// Mostrar/Ocultar tarjeta explicativa de carga de Excel
function mostrarEstadoSinDatos(mostrar) {
    const noDataCard = document.getElementById("no-data-state");
    const kpiGrid = document.getElementById("kpi-grid");
    const chartsLayout = document.getElementById("charts-layout");
    const detailTableCard = document.getElementById("detail-table-card");

    if (mostrar) {
        noDataCard.classList.remove("hidden");
        kpiGrid.classList.add("hidden");
        chartsLayout.classList.add("hidden");
        detailTableCard.classList.add("hidden");
    } else {
        noDataCard.classList.add("hidden");
        kpiGrid.classList.remove("hidden");
        chartsLayout.classList.remove("hidden");
        detailTableCard.classList.remove("hidden");
    }
}
