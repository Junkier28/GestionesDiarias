// Configuración global y funciones comunes
const config = {
    jerarquias: {
        dmiro: [
            "NUMERO EQUIVOCADO", "NO VIVE EN DOMICILIO", "NO SE ENCUENTRA EN DOMICILIO",
            "LINEA AVERIADA / LINEA OCUPADA", "CONTESTADOR AUTOMATICO", "CUELGA LLAMADA",
            "NO CONTESTA", "MENSAJE A TERCEROS", "NOTIFICACIÓN FAMILIARES / TERCEROS",
            "WHATSAPP ENVIADO POR GESTOR", "NO AYUDA", "NEGATIVA DE PAGO",
            "INTERESADO NEGOCIACIÓN", "VOLVER A LLAMAR", "POSTERGA PAGO",
            "ACUERDO DE PAGO / CONVENIO", "ACUERDO DE PAGO / VERBAL", "ABONO PARCIAL",
            "RECORDATORIO DE PROMESA", "PAGO REALIZADO", "LIQUIDACIÓN"
        ],
        pacifico_propia: [
            "NUMERO EQUIVOCADO", "LINEA AVERIADA / LINEA OCUPADA", "CONTESTADOR AUTOMATICO",
            "CUELGA LLAMADA", "NO CONTESTA", "MENSAJE A TERCEROS", "MENSAJE FAMILIAR",
            "CORREO ENVIADO POR GESTOR", "WHATSAPP ENVIADO POR GESTOR",
            "WHATSAPP RECIBIDO POR CLIENTE", "NO AYUDA", "NEGATIVA DE PAGO",
            "INTERESADO EN NEGOCIACION", "INTERESADO NEGOCIACIÓN", "VOLVER A LLAMAR",
            "POSTERGA PAGO", "PROMESA DE PAGO - ABONO PARCIAL",
            "PROMESA DE PAGO - LIQUIDACIÓN", "PROMESA DE PAGO ACUERDO DE PAGO / VERBAL",
            "RECORDATORIO DE PROMESA", "PAGO REALIZADO", "CLIENTE YA LIQUIDÓ DEUDA"
        ],
        pacifico_servicio: [
            "NUMERO EQUIVOCADO", "ILOCALIZABLE TELLEFONICAMENTE", "LINEA AVERIADA / LINEA OCUPADA",
            "CONTESTADOR AUTOMATICO", "CUELGA LLAMADA", "NO CONTESTA", "FALLECIDO",
            "MENSAJE A TERCEROS", "MENSAJE FAMILIAR", "CORREO ENVIADO POR GESTOR",
            "WHATSAPP ENVIADO POR GESTOR", "NO AYUDA", "NEGATIVA DE PAGO",
            "INTERESADO EN NEGOCIACION", "INTERESADO NEGOCIACION", "VOLVER A LLAMAR",
            "PROMESA DE PAGO - ABONO PARCIAL", "PROMESA DE PAGO - LIQUIDACION",
            "PROMESA DE PAGO ACUERDO DE PAGO / VERBAL",
            "PROMESA DE PAGO ACUERDO DE PAGO / ACTA TRANSACCIONAL",
            "RECORDATORIO DE PROMESA", "PAGO REALIZADO", "CLIENTE YA LIQUIDO DEUDA"
        ]
    }
};

// Variables globales
let processedFiles = [];
let selectedOption = 'mejor_respuesta';

// Elementos del DOM
const uploadArea = document.getElementById('uploadArea');
const fileInput = document.getElementById('fileInput');
const progressContainer = document.getElementById('progressContainer');
const progressFill = document.getElementById('progressFill');
const status = document.getElementById('status');
const results = document.getElementById('results');
const resultsContent = document.getElementById('resultsContent');
const llamadasDiaContainer = document.getElementById('llamadasDiaContainer');
const exportLlamadasExcel = document.getElementById('exportLlamadasExcel');
const resetLlamadasReport = document.getElementById('resetLlamadasReport');

// Inicialización
document.addEventListener('DOMContentLoaded', () => {
    setupEventListeners();
});

// Configurar event listeners
function setupEventListeners() {
    // Área de carga - SOLUCIÓN: No abrir input al hacer click si ya hay archivos
    uploadArea.addEventListener('click', (e) => {
        // Evitar que el click en el input file propague al área
        if (e.target === fileInput) {
            return;
        }
        
        // Solo abrir si no estamos en un botón de descarga u otro elemento
        if (e.target.classList.contains('download-btn') || 
            e.target.closest('.download-btn') ||
            e.target.classList.contains('export-btn')) {
            return;
        }
        
        // Limpiar el input file antes de abrir
        fileInput.value = '';
        fileInput.click();
    });
    
    // Prevenir que el click del input file se propague
    fileInput.addEventListener('click', (e) => {
        e.stopPropagation();
    });
    
    fileInput.addEventListener('change', handleFiles);
    
    // Botones de opciones
    document.querySelectorAll('.option-btn').forEach(btn => {
        btn.addEventListener('click', () => {
            document.querySelectorAll('.option-btn').forEach(b => b.classList.remove('active'));
            btn.classList.add('active');
            selectedOption = btn.dataset.option;
            
            // Ocultar contenedores específicos
            llamadasDiaContainer.style.display = 'none';
            results.style.display = 'none';
            resultsContent.innerHTML = '';
            processedFiles = [];
            
            // Limpiar input file
            fileInput.value = '';
        });
    });
    
    // Drag and drop
    uploadArea.addEventListener('dragover', (e) => {
        e.preventDefault();
        uploadArea.classList.add('dragover');
    });
    
    uploadArea.addEventListener('dragleave', () => {
        uploadArea.classList.remove('dragover');
    });
    
    uploadArea.addEventListener('drop', (e) => {
        e.preventDefault();
        uploadArea.classList.remove('dragover');
        
        // Limpiar archivos anteriores
        processedFiles = [];
        fileInput.value = '';
        
        handleFiles({
            target: {
                files: e.dataTransfer.files
            }
        });
    });
    
    // Botones del reporte de llamadas
    if (exportLlamadasExcel) {
        exportLlamadasExcel.addEventListener('click', () => {
            const file = processedFiles.find(f => f.opcion === 'llamadas_dia');
            if (file) {
                exportarLlamadasDia(file.datos, file.nombre);
            }
        });
    }
    
    if (resetLlamadasReport) {
        resetLlamadasReport.addEventListener('click', resetearReporteLlamadas);
    }
}

// Detectar tipo de archivo
function detectarTipo(headers, filaEncabezado) {
    const columnas = Object.keys(headers);
    
    // DMIRO - Fila 5 con columnas específicas
    if (filaEncabezado === 5 && columnas.includes('Documento') && columnas.includes('Nivel 2')) {
        return 'dmiro';
    }
    
    // PACÍFICO - Fila 3
    if (filaEncabezado === 3 && columnas.includes('Número Credito') && columnas.includes('Nivel 3')) {
        const tieneIdPromesa = columnas.includes('Fecha Promesa');
        return tieneIdPromesa ? 'pacifico_propia' : 'pacifico_servicio';
    }
    
    return null;
}

// Leer archivo Excel
async function leerArchivoExcel(file) {
    const arrayBuffer = await file.arrayBuffer();
    const workbook = XLSX.read(arrayBuffer, { type: 'array' });
    return workbook;
}

// Procesar archivo según opción
function procesarArchivo(data, tipo, opcion) {
    switch(opcion) {
        case 'mejor_respuesta':
            return procesarMejorRespuesta(data, tipo);
        case 'numero_gestiones':
            return procesarNumeroGestiones(data, tipo);
        case 'tiempo_muerto':
            return procesarTiempoMuerto(data, tipo);
        case 'consolidar_comentarios':
            return procesarConsolidarComentarios(data, tipo);
        case 'llamadas_dia':
            return procesarLlamadasDia(data, tipo);
        default:
            return procesarMejorRespuesta(data, tipo);
    }
}

// Manejar archivos cargados
async function handleFiles(event) {
    const files = Array.from(event.target.files);
    if (files.length === 0) return;
    
    // Limpiar resultados anteriores
    processedFiles = [];
    results.style.display = 'none';
    resultsContent.innerHTML = '';
    llamadasDiaContainer.style.display = 'none';
    
    showStatus('Iniciando procesamiento...', 'info');
    updateProgress(0);
    
    // Si es reporte de llamadas por día, solo procesar un archivo
    const esLlamadasDia = selectedOption === 'llamadas_dia';
    const archivosAProcesar = esLlamadasDia ? [files[0]] : files;
    
    for (let i = 0; i < archivosAProcesar.length; i++) {
        const file = archivosAProcesar[i];
        showStatus(`Procesando ${file.name} (${i + 1}/${archivosAProcesar.length})`, 'info');
        
        try {
            let workbook, data, tipo = null;
            
            // Probar diferentes filas de encabezado
            for (const filaEnc of [3, 5]) {
                try {
                    workbook = await leerArchivoExcel(file);
                    const worksheet = workbook.Sheets[workbook.SheetNames[0]];
                    
                    // Convertir a JSON con encabezado específico
                    const range = XLSX.utils.decode_range(worksheet['!ref']);
                    range.s.r = filaEnc - 1;
                    
                    const newRef = XLSX.utils.encode_range(range);
                    const tempWs = {
                        ...worksheet,
                        '!ref': newRef
                    };
                    
                    data = XLSX.utils.sheet_to_json(tempWs);
                    
                    if (data.length > 0) {
                        tipo = detectarTipo(data[0], filaEnc);
                        if (tipo) break;
                    }
                } catch (e) {
                    console.warn(`No se pudo procesar con fila ${filaEnc}:`, e);
                    continue;
                }
            }
            
            if (!tipo || !data || data.length === 0) {
                throw new Error('No se pudo identificar el tipo de archivo');
            }
            
            const datosProcesados = procesarArchivo(data, tipo, selectedOption);
            
            // Para Reporte Llamadas/Día, mostrar directamente
            if (esLlamadasDia) {
                mostrarReporteLlamadas(datosProcesados, file.name);
                processedFiles.push({
                    nombre: file.name,
                    tipo: tipo,
                    opcion: selectedOption,
                    registrosOriginales: data.length,
                    datos: datosProcesados
                });
                break;
            } else {
                // Para otras opciones
                const nombreOriginal = file.name.replace(/\.[^/.]+$/, "");
                const nombreSalida = `${nombreOriginal}_${selectedOption}.xlsx`;
                
                let registrosProcesados = 0;
                if (datosProcesados.tipo === 'numero_gestiones') {
                    registrosProcesados = Object.keys(datosProcesados.tablaDinamica).length;
                } else if (datosProcesados.tipo === 'tiempo_muerto') {
                    registrosProcesados = datosProcesados.resumenTiempoMuerto.length;
                } else {
                    registrosProcesados = datosProcesados.length || 0;
                }
                
                processedFiles.push({
                    nombre: file.name,
                    tipo: tipo,
                    opcion: selectedOption,
                    registrosOriginales: data.length,
                    registrosProcesados: registrosProcesados,
                    datos: datosProcesados,
                    nombreSalida: nombreSalida
                });
            }
            
        } catch (error) {
            console.error('Error procesando archivo:', error);
            processedFiles.push({
                nombre: file.name,
                error: error.message
            });
        }
        
        updateProgress(((i + 1) / archivosAProcesar.length) * 100);
    }
    
    // Limpiar input file después de procesar
    fileInput.value = '';
    
    if (!esLlamadasDia) {
        mostrarResultados();
    }
    
    showStatus('¡Procesamiento completado!', 'success');
    setTimeout(() => {
        updateProgress(0);
        showStatus('', 'info');
    }, 2000);
}

// Mostrar resultados generales
function mostrarResultados() {
    if (processedFiles.length === 0) return;
    
    resultsContent.innerHTML = '';
    results.style.display = 'block';
    
    // Filtrar archivos sin error
    const archivosExitosos = processedFiles.filter(f => !f.error);
    
    if (archivosExitosos.length === 0) {
        resultsContent.innerHTML = `
            <div class="result-item">
                <h4>❌ Error en todos los archivos</h4>
                <p style="color: #c62828;">No se pudo procesar ningún archivo correctamente.</p>
            </div>
        `;
        return;
    }
    
    archivosExitosos.forEach(file => {
        let registrosInfo = '';
        
        if (file.opcion === 'numero_gestiones') {
            const numAgentes = Object.keys(file.datos.tablaDinamica).length;
            registrosInfo = `
                <div class="stat">
                    <span class="stat-number">${numAgentes}</span>
                    <span class="stat-label">Agentes</span>
                </div>
                <div class="stat">
                    <span class="stat-number">${file.datos.totales['Total general']}</span>
                    <span class="stat-label">Total Gestiones</span>
                </div>
            `;
        } else if (file.opcion === 'tiempo_muerto') {
            const numAgentes = file.datos.resumenTiempoMuerto.length;
            registrosInfo = `
                <div class="stat">
                    <span class="stat-number">${numAgentes}</span>
                    <span class="stat-label">Agentes</span>
                </div>
            `;
        } else {
            registrosInfo = `
                <div class="stat">
                    <span class="stat-number">${file.registrosProcesados}</span>
                    <span class="stat-label">Registros</span>
                </div>
            `;
        }
        
        resultsContent.innerHTML += `
            <div class="result-item">
                <h4>✅ ${file.nombre}</h4>
                <div class="result-stats">
                    <div class="stat">
                        <span class="stat-number">${file.tipo}</span>
                        <span class="stat-label">Tipo Detectado</span>
                    </div>
                    <div class="stat">
                        <span class="stat-number">${file.registrosOriginales}</span>
                        <span class="stat-label">Registros Originales</span>
                    </div>
                    ${registrosInfo}
                </div>
                <button class="download-btn" onclick="descargarArchivo('${file.nombre}')">
                    📥 Descargar Procesado
                </button>
            </div>
        `;
    });
    
    // Mostrar errores si los hay
    const archivosConError = processedFiles.filter(f => f.error);
    archivosConError.forEach(file => {
        resultsContent.innerHTML += `
            <div class="result-item">
                <h4>❌ ${file.nombre}</h4>
                <p style="color: #c62828;">Error: ${file.error}</p>
            </div>
        `;
    });
}

// Descargar archivo procesado
async function descargarArchivo(nombreOriginal) {
    const file = processedFiles.find(f => f.nombre === nombreOriginal);
    if (!file || !file.datos) {
        showStatus('No se encontraron datos para descargar', 'error');
        return;
    }
    
    // Convertir nombre de opción a formato CamelCase
    const opcionFormateada = file.opcion
        .split('_')
        .map(word => word.charAt(0).toUpperCase() + word.slice(1))
        .join('');
    
    const funcionExportar = window[`exportar${opcionFormateada}`];
    
    if (typeof funcionExportar === 'function') {
        try {
            await funcionExportar(file.datos, file.nombreSalida);
            showStatus(`Archivo "${file.nombreSalida}" descargado`, 'success');
        } catch (error) {
            console.error('Error al exportar:', error);
            showStatus('Error al generar el archivo de descarga', 'error');
        }
    } else {
        // Fallback a exportación básica
        descargarExcelBasico(file.datos, file.nombreSalida);
    }
}

// Descargar Excel básico
function descargarExcelBasico(datos, nombreArchivo) {
    try {
        const wb = XLSX.utils.book_new();
        const ws = XLSX.utils.json_to_sheet(datos);
        XLSX.utils.book_append_sheet(wb, ws, "Procesado");
        XLSX.writeFile(wb, nombreArchivo);
        showStatus(`Archivo "${nombreArchivo}" descargado`, 'success');
    } catch (error) {
        console.error('Error al exportar Excel básico:', error);
        showStatus('Error al generar el archivo de descarga', 'error');
    }
}

// Mostrar estado
function showStatus(message, type = 'info') {
    status.textContent = message;
    status.className = `status ${type}`;
    status.style.display = 'block';
}

// Actualizar progreso
function updateProgress(percent) {
    progressFill.style.width = `${percent}%`;
    progressContainer.style.display = percent === 0 ? 'none' : 'block';
}

// Funciones de utilidad
function formatearTiempo(segundos) {
    const horas = Math.floor(segundos / 3600);
    const minutos = Math.floor((segundos % 3600) / 60);
    const segs = segundos % 60;
    return `${horas.toString().padStart(2, '0')}:${minutos.toString().padStart(2, '0')}:${segs.toString().padStart(2, '0')}`;
}

function parsearFechaExcel(valor) {
    if (!valor) return null;
    
    // Si es número de Excel
    if (typeof valor === 'number') {
        const excelEpoch = new Date(1899, 11, 30);
        const fecha = new Date(excelEpoch.getTime() + valor * 24 * 60 * 60 * 1000);
        return fecha;
    }
    
    // Si es string
    if (typeof valor === 'string') {
        // Intentar como fecha ISO
        const fechaISO = new Date(valor);
        if (!isNaN(fechaISO.getTime())) return fechaISO;
        
        // Intentar diferentes formatos
        const formatos = [
            { regex: /(\d{2})\/(\d{2})\/(\d{4})/, indices: [2, 1, 0] }, // DD/MM/YYYY
            { regex: /(\d{4})-(\d{2})-(\d{2})/, indices: [0, 1, 2] }, // YYYY-MM-DD
            { regex: /(\d{2})-(\d{2})-(\d{4})/, indices: [2, 1, 0] }, // DD-MM-YYYY
            { regex: /(\d{1,2})\/(\d{1,2})\/(\d{4}) (\d{1,2}):(\d{2}):(\d{2})/, indices: [2, 1, 0] } // DD/MM/YYYY HH:MM:SS
        ];
        
        for (const formato of formatos) {
            const match = valor.match(formato.regex);
            if (match) {
                const dia = parseInt(match[formato.indices[0] + 1]);
                const mes = parseInt(match[formato.indices[1] + 1]) - 1;
                const anio = parseInt(match[formato.indices[2] + 1]);
                const fecha = new Date(anio, mes, dia);
                
                if (match.length > 3 && formato.indices.length > 3) {
                    // Si hay hora
                    const hora = parseInt(match[formato.indices[3] + 1]) || 0;
                    const minutos = parseInt(match[formato.indices[4] + 1]) || 0;
                    const segundos = parseInt(match[formato.indices[5] + 1]) || 0;
                    fecha.setHours(hora, minutos, segundos);
                }
                
                if (!isNaN(fecha.getTime())) {
                    return fecha;
                }
            }
        }
    }
    
    // Si es objeto Date
    if (valor instanceof Date) {
        return valor;
    }
    
    return null;
}

// Función para mostrar reporte de llamadas (se implementará en llamadas-dia.js)
function mostrarReporteLlamadas(datos, nombreArchivo) {
    // Esta función será implementada en llamadas-dia.js
    if (window.mostrarReporteLlamadasCompleto) {
        window.mostrarReporteLlamadasCompleto(datos, nombreArchivo);
    }
}

// Función para resetear reporte (se implementará en llamadas-dia.js)
function resetearReporteLlamadas() {
    // Esta función será implementada en llamadas-dia.js
    if (window.resetearReporteLlamadasCompleto) {
        window.resetearReporteLlamadasCompleto();
    } else {
        // Fallback
        llamadasDiaContainer.style.display = 'none';
        processedFiles = [];
        fileInput.value = '';
        showStatus('', 'info');
    }
}

// Exportar funciones al scope global
window.descargarArchivo = descargarArchivo;
window.showStatus = showStatus;
window.updateProgress = updateProgress;
window.parsearFechaExcel = parsearFechaExcel;
window.formatearTiempo = formatearTiempo;