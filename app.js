// Variables globales para almacenar los datos de la aplicación y las instancias de gráficos
let allData = [];
let filteredData = [];
let charts = {
    physical: null,
    certified: null,
    accrued: null,
    reprogramming: null
};

// Configuración de la carga inicial del archivo Excel
document.addEventListener("DOMContentLoaded", () => {
    inicializarEventos();
    
    // ALTERNATIVA B: Carga automática de "data.xlsx" en la raíz del proyecto
    // Para activar la ALTERNATIVA A (solo carga manual), comente o borre el bloque de código abajo
    // y elimine la clase 'hidden' del contenedor de estado #no-data-state en index.html
    cargarExcelAutomatico("data.xlsx");
});

// Función para inicializar los event listeners de los filtros y el input de archivo
function inicializarEventos() {
    // Input de archivo para Alternativa A
    const fileInput = document.getElementById("excel-file-input");
    fileInput.addEventListener("change", cargarExcelManual);

    // Eventos para filtros en cascada
    const filters = ["filter-cod-proy", "filter-cite", "filter-pet", "filter-reprog"];
    filters.forEach(id => {
        document.getElementById(id).addEventListener("change", () => {
            aplicarFiltros();
        });
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
            const financ_deven = parseCleanNumber(row["financ_deven"]);

            // 4. Limpieza y transformación de porcentajes
            const fisica_prog = parseCleanNumber(row["fisica_prog"]);
            const fisica_ejec = parseCleanNumber(row["fisica_ejec"]);
            
            const fisica_pct = parsePercentage(row["fisica_pct"]);
            const financ_cert_pct = parsePercentage(row["financ_cert_pct"]);
            const financ_deven_pct = parsePercentage(row["financ_deven_pct"]);

            // 5. Normalizar "req_reprogr"
            let req_reprogr = (row["req_reprogr"] || "No").toString().trim();
            if (/^s[ií]/i.test(req_reprogr)) {
                req_reprogr = "Sí";
            } else {
                req_reprogr = "No";
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
                financ_deven: financ_deven,
                financ_cert_pct: financ_cert_pct,
                financ_deven_pct: financ_deven_pct,
                req_reprogr: req_reprogr
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
    const selCod = document.getElementById("filter-cod-proy");
    const selCite = document.getElementById("filter-cite");
    const selPet = document.getElementById("filter-pet");
    const selReprog = document.getElementById("filter-reprog");

    // Guardar los valores previamente seleccionados
    const valCod = selCod.value;
    const valCite = selCite.value;
    const valPet = selPet.value;
    const valReprog = selReprog.value;

    // Generar opciones filtrando los datos del resto de componentes
    
    // Códigos de proyecto disponibles considerando los otros filtros
    const dataForCod = allData.filter(d => 
        (!valCite || d.cite === valCite) &&
        (!valPet || d.pet === valPet) &&
        (!valReprog || d.req_reprogr === valReprog)
    );
    const optCod = [...new Set(dataForCod.map(d => d.codigo_proyecto))].sort();

    // CITEs disponibles considerando los otros filtros
    const dataForCite = allData.filter(d => 
        (!valCod || d.codigo_proyecto === valCod) &&
        (!valPet || d.pet === valPet) &&
        (!valReprog || d.req_reprogr === valReprog)
    );
    const optCite = [...new Set(dataForCite.map(d => d.cite))].sort();

    // PETs disponibles considerando los otros filtros
    const dataForPet = allData.filter(d => 
        (!valCod || d.codigo_proyecto === valCod) &&
        (!valCite || d.cite === valCite) &&
        (!valReprog || d.req_reprogr === valReprog)
    );
    const optPet = [...new Set(dataForPet.map(d => d.pet))].sort();

    // Reprogramaciones disponibles considerando los otros filtros
    const dataForReprog = allData.filter(d => 
        (!valCod || d.codigo_proyecto === valCod) &&
        (!valCite || d.cite === valCite) &&
        (!valPet || d.pet === valPet)
    );
    const optReprog = [...new Set(dataForReprog.map(d => d.req_reprogr))].sort();

    // Volver a rellenar los selects preservando los valores
    rellenarSelect(selCod, optCod, valCod, "Todos los códigos");
    rellenarSelect(selCite, optCite, valCite, "Todos los CITEs");
    rellenarSelect(selPet, optPet, valPet, "Todos los PET");
    rellenarSelect(selReprog, optReprog, valReprog, "Todos");
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
    const valCod = document.getElementById("filter-cod-proy").value;
    const valCite = document.getElementById("filter-cite").value;
    const valPet = document.getElementById("filter-pet").value;
    const valReprog = document.getElementById("filter-reprog").value;

    // Filtrar el conjunto de datos completo
    filteredData = allData.filter(d => {
        return (!valCod || d.codigo_proyecto === valCod) &&
               (!valCite || d.cite === valCite) &&
               (!valPet || d.pet === valPet) &&
               (!valReprog || d.req_reprogr === valReprog);
    });

    // Actualizar el estado de los filtros en cascada
    actualizarFiltrosCascada();
    
    // Refrescar KPIs, Gráficos y Tabla
    actualizarDashboard();
}

// Limpiar filtros restableciendo todos los valores y datos
function limpiarFiltros() {
    document.getElementById("filter-cod-proy").value = "";
    document.getElementById("filter-cite").value = "";
    document.getElementById("filter-pet").value = "";
    document.getElementById("filter-reprog").value = "";

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
    let citesOver50 = 0;
    let totalProg = 0;
    let totalCert = 0;
    let totalDeven = 0;
    let totalReprog = 0;
    let pctCert = 0;
    let pctDeven = 0;

    if (totalProy > 0) {
        // CITEs con física > 50%
        const citeData = agruparPorCite(filteredData);
        citesOver50 = citeData.filter(c => c.fisica_pct > 50).length;

        // Sumas de montos
        totalProg = filteredData.reduce((acc, curr) => acc + curr.financ_prog, 0);
        totalCert = filteredData.reduce((acc, curr) => acc + curr.financ_cert, 0);
        totalDeven = filteredData.reduce((acc, curr) => acc + curr.financ_deven, 0);

        // Reprogramaciones
        totalReprog = filteredData.filter(d => d.req_reprogr === "Sí").length;

        // Porcentajes globales
        pctCert = totalProg > 0 ? (totalCert / totalProg) * 100 : 0;
        pctDeven = totalProg > 0 ? (totalDeven / totalProg) * 100 : 0;
    }

    document.getElementById("kpi-val-projects").textContent = totalProy;
    document.getElementById("kpi-val-cites-50").textContent = citesOver50;
    document.getElementById("kpi-val-prog").textContent = formatearSoles(totalProg);
    document.getElementById("kpi-val-cert").textContent = formatearSoles(totalCert);
    document.getElementById("kpi-val-cert-pct").textContent = `${pctCert.toFixed(1)}%`;
    document.getElementById("kpi-val-deven").textContent = formatearSoles(totalDeven);
    document.getElementById("kpi-val-deven-pct").textContent = `${pctDeven.toFixed(1)}%`;
    document.getElementById("kpi-val-reprog").textContent = totalReprog;
}

// Renderizar gráficos de Chart.js
function actualizarGraficos() {
    destruirGraficosExistentes();

    if (filteredData.length === 0) return;

    // Preparar datos ordenados para gráficos horizontales por CITE/Código (usando CITE)
    // Agrupamos por CITE para mejor visualización de avance si hay duplicados
    const citeData = agruparPorCite(filteredData);

    // Gráfico 1: % Ejecución Física (Barras Horizontales - Ordenado de mayor a menor)
    const sortedPhysical = [...citeData].sort((a, b) => b.fisica_pct - a.fisica_pct);
    const labelsPhysical = sortedPhysical.map(d => d.cite);
    const dataPhysical = sortedPhysical.map(d => d.fisica_pct);
    
    // Colores semáforo para el gráfico físico:
    // Rojo < 25%, Amarillo [25%, 50%], Verde > 50%
    const colorsPhysical = dataPhysical.map(val => {
        if (val < 25) return "rgba(239, 68, 68, 0.85)";       // Rojo
        if (val <= 50) return "rgba(245, 158, 11, 0.85)";     // Amarillo/Naranja
        return "rgba(16, 185, 129, 0.85)";                    // Verde
    });
    const bordersPhysical = dataPhysical.map(val => {
        if (val < 25) return "rgba(239, 68, 68, 1)";
        if (val <= 50) return "rgba(245, 158, 11, 1)";
        return "rgba(16, 185, 129, 1)";
    });

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
        options: obtenerConfiguracionGraficoBarras(true) // true para horizontal
    });

    // Gráfico 2: % Certificado (Barras Horizontales - Ordenado de mayor a menor)
    const sortedCertified = [...citeData].sort((a, b) => b.financ_cert_pct - a.financ_cert_pct);
    const labelsCertified = sortedCertified.map(d => d.cite);
    const dataCertified = sortedCertified.map(d => d.financ_cert_pct);

    const ctxCertified = document.getElementById("chart-certified").getContext("2d");
    charts.certified = new Chart(ctxCertified, {
        type: "bar",
        data: {
            labels: labelsCertified,
            datasets: [{
                label: "% Certificado Financiero",
                data: dataCertified,
                backgroundColor: "rgba(180, 138, 39, 0.85)", // Dorado suave
                borderColor: "rgba(180, 138, 39, 1)",
                borderWidth: 1.5,
                borderRadius: 4
            }]
        },
        options: obtenerConfiguracionGraficoBarras(true)
    });

    // Gráfico 3: % Devengado (Barras Horizontales - Ordenado de mayor a menor)
    const sortedAccrued = [...citeData].sort((a, b) => b.financ_deven_pct - a.financ_deven_pct);
    const labelsAccrued = sortedAccrued.map(d => d.cite);
    const dataAccrued = sortedAccrued.map(d => d.financ_deven_pct);

    const ctxAccrued = document.getElementById("chart-accrued").getContext("2d");
    charts.accrued = new Chart(ctxAccrued, {
        type: "bar",
        data: {
            labels: labelsAccrued,
            datasets: [{
                label: "% Devengado Financiero",
                data: dataAccrued,
                backgroundColor: "rgba(15, 23, 42, 0.85)", // Azul marino oscuro
                borderColor: "rgba(15, 23, 42, 1)",
                borderWidth: 1.5,
                borderRadius: 4
            }]
        },
        options: obtenerConfiguracionGraficoBarras(true)
    });

    // Gráfico 4: Torta/Dona de Reprogramación (Sí/No)
    let conteoSí = 0;
    let conteoNo = 0;
    filteredData.forEach(d => {
        if (d.req_reprogr === "Sí") conteoSí++;
        else conteoNo++;
    });

    const ctxReprog = document.getElementById("chart-reprogramming").getContext("2d");
    charts.reprogramming = new Chart(ctxReprog, {
        type: "doughnut",
        data: {
            labels: ["Requiere Reprogramación (Sí)", "No Requiere (No)"],
            datasets: [{
                data: [conteoSí, conteoNo],
                backgroundColor: [
                    "rgba(180, 138, 39, 0.85)", // Oro/Dorado
                    "rgba(15, 23, 42, 0.85)"   // Azul marino
                ],
                borderColor: [
                    "rgba(180, 138, 39, 1)",
                    "rgba(15, 23, 42, 1)"
                ],
                borderWidth: 1.5
            }]
        },
        options: {
            responsive: true,
            maintainAspectRatio: false,
            plugins: {
                legend: {
                    position: "bottom",
                    labels: {
                        font: { family: "Inter", size: 11 },
                        color: "#475569"
                    }
                },
                tooltip: {
                    callbacks: {
                        label: function(context) {
                            const val = context.raw;
                            const total = conteoSí + conteoNo;
                            const pct = total > 0 ? ((val / total) * 100).toFixed(1) : 0;
                            return ` ${context.label}: ${val} (${pct}%)`;
                        }
                    }
                }
            },
            cutout: "60%"
        }
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
                financ_cert_pct_sum: 0,
                financ_deven_pct_sum: 0,
                count: 0
            };
        }
        groups[item.cite].fisica_pct_sum += item.fisica_pct;
        groups[item.cite].financ_cert_pct_sum += item.financ_cert_pct;
        groups[item.cite].financ_deven_pct_sum += item.financ_deven_pct;
        groups[item.cite].count++;
    });

    return Object.values(groups).map(g => ({
        cite: g.cite,
        fisica_pct: g.fisica_pct_sum / g.count,
        financ_cert_pct: g.financ_cert_pct_sum / g.count,
        financ_deven_pct: g.financ_deven_pct_sum / g.count
    }));
}

// Configuración común de ejes y leyendas para los gráficos de barras
function obtenerConfiguracionGraficoBarras(isHorizontal) {
    return {
        indexAxis: isHorizontal ? "y" : "x",
        responsive: true,
        maintainAspectRatio: false,
        plugins: {
            legend: {
                display: false
            },
            tooltip: {
                callbacks: {
                    label: function(context) {
                        return ` Avance: ${context.raw.toFixed(1)}%`;
                    }
                }
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
                    color: "#475569"
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
        
        // Crear celdas con heatmap y estilos dinámicos
        // Para el heatmap usamos opacidad de colores corporativos
        const styleFisico = `background-color: rgba(30, 58, 138, ${row.fisica_pct / 100 * 0.22}); font-weight: 500;`;
        const styleCertificado = `background-color: rgba(180, 138, 39, ${row.financ_cert_pct / 100 * 0.22}); font-weight: 500;`;
        const styleDevengado = `background-color: rgba(15, 23, 42, ${row.financ_deven_pct / 100 * 0.22}); font-weight: 500;`;

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
            <td class="text-right" style="white-space: nowrap;">${formatearSoles(row.financ_deven)}</td>
            <td class="text-right" style="${styleCertificado}">${row.financ_cert_pct.toFixed(1)}%</td>
            <td class="text-right" style="${styleDevengado}">${row.financ_deven_pct.toFixed(1)}%</td>
            <td class="text-center">
                <span class="badge ${row.req_reprogr === "Sí" ? "btn-primary" : "badge-light"}" style="font-size: 0.7rem; padding: 0.15rem 0.45rem;">
                    ${row.req_reprogr}
                </span>
            </td>
        `;
        
        tableBody.appendChild(tr);
    });
}

// Calcular Alertas de acuerdo con la lógica operativa-financiera
function calcularAlerta(row) {
    const fisica = row.fisica_pct;
    const cert = row.financ_cert_pct;
    const deven = row.financ_deven_pct;

    // Lógica 1: Crítico operativo (fisica_pct < 25 y financ_deven_pct < 25)
    if (fisica < 25 && deven < 25) {
        return {
            texto: "Crítico operativo",
            clase: "alert-critico"
        };
    }

    // Lógica 2: Riesgo financiero (financ_cert_pct > 50 y financ_deven_pct < 25)
    if (cert > 50 && deven < 25) {
        return {
            texto: "Riesgo financiero",
            clase: "alert-riesgo"
        };
    }

    // Lógica 3: Desfase físico-financiero (financ_cert_pct - fisica_pct > 30)
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
