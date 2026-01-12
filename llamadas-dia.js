// Procesador de Reporte Llamadas por Día
function procesarLlamadasDia(data, tipo) {
    const fechaColumna = tipo === 'dmiro' ? 'Fecha' : 'Fecha de Gestión';
    const agenteColumna = 'Agente';
    
    const agentesData = {};
    const fechasSet = new Set();
    const todosRegistros = [];
    
    // Procesar cada registro
    data.forEach(row => {
        const fechaRaw = row[fechaColumna];
        const agente = row[agenteColumna];
        
        if (!fechaRaw || !agente) return;
        
        // Parsear fecha
        const fechaObj = parsearFechaExcel(fechaRaw);
        if (!fechaObj || isNaN(fechaObj.getTime())) return;
        
        const fechaStr = fechaObj.toISOString().split('T')[0];
        const agenteNombre = agente.toString().trim();
        
        // Inicializar datos del agente
        if (!agentesData[agenteNombre]) {
            agentesData[agenteNombre] = {
                diario: {},
                total: 0
            };
        }
        
        // Inicializar contador del día
        if (!agentesData[agenteNombre].diario[fechaStr]) {
            agentesData[agenteNombre].diario[fechaStr] = 0;
        }
        
        // Incrementar contadores
        agentesData[agenteNombre].diario[fechaStr]++;
        agentesData[agenteNombre].total++;
        
        // Agregar fecha al conjunto
        fechasSet.add(fechaStr);
        
        // Guardar registro completo
        todosRegistros.push({
            fecha: fechaStr,
            agente: agenteNombre,
            registroCompleto: row
        });
    });
    
    return {
        tipo: 'llamadas_dia',
        agentesData: agentesData,
        fechas: Array.from(fechasSet).sort(),
        todosRegistros: todosRegistros,
        registrosProcesados: todosRegistros.length
    };
}

// Mostrar reporte de llamadas por día
function mostrarReporteLlamadasCompleto(datos, nombreArchivo) {
    // Ocultar resultados generales
    results.style.display = 'none';
    
    // Mostrar contenedor específico
    llamadasDiaContainer.style.display = 'block';
    
    const agentesData = datos.agentesData;
    const fechas = datos.fechas;
    const agentes = Object.keys(agentesData).sort();
    
    // Actualizar estadísticas
    actualizarEstadisticasLlamadas(agentesData, fechas, datos.todosRegistros.length);
    
    // Generar tabla
    generarTablaLlamadas(agentesData, fechas, agentes);
    
    // Actualizar nombre del archivo en el botón de exportar
    const exportBtn = document.getElementById('exportLlamadasExcel');
    if (exportBtn) {
        exportBtn.dataset.nombreArchivo = nombreArchivo;
        exportBtn.dataset.datos = JSON.stringify(datos);
    }
}

// Actualizar estadísticas
function actualizarEstadisticasLlamadas(agentesData, fechas, totalLlamadas) {
    document.getElementById('totalCalls').textContent = totalLlamadas.toLocaleString();
    document.getElementById('uniqueAgents').textContent = Object.keys(agentesData).length.toLocaleString();
    
    if (fechas.length > 0) {
        const primeraFecha = new Date(fechas[0]);
        const ultimaFecha = new Date(fechas[fechas.length - 1]);
        
        const formatoFecha = (fecha) => fecha.toLocaleDateString('es-ES', {
            day: '2-digit',
            month: '2-digit',
            year: 'numeric'
        });
        
        document.getElementById('dateRange').textContent = 
            `${formatoFecha(primeraFecha)} - ${formatoFecha(ultimaFecha)}`;
    }
    
    // Top 3 agentes
    const topAgentes = Object.keys(agentesData)
        .map(agente => ({
            nombre: agente,
            total: agentesData[agente].total
        }))
        .sort((a, b) => b.total - a.total)
        .slice(0, 3);
    
    const topAgentsDiv = document.getElementById('topAgents');
    topAgentsDiv.innerHTML = '';
    
    if (topAgentes.length === 0) {
        topAgentsDiv.innerHTML = '<div style="color: #7f8c8d; text-align: center;">No hay datos de agentes</div>';
    } else {
        topAgentes.forEach((agente, index) => {
            const medallas = ['🥇', '🥈', '🥉'];
            const div = document.createElement('div');
            div.innerHTML = `
                <span>${medallas[index]} ${agente.nombre}</span>
                <span style="background: #3498db; color: white; padding: 5px 10px; border-radius: 20px; font-size: 0.9em; font-weight: 600;">
                    ${agente.total} llamadas
                </span>
            `;
            topAgentsDiv.appendChild(div);
        });
    }
}

// Generar tabla HTML
function generarTablaLlamadas(agentesData, fechas, agentes) {
    const thead = document.getElementById('llamadasTableHeader');
    const tbody = document.getElementById('llamadasTableBody');
    
    // Limpiar tabla anterior
    thead.innerHTML = '';
    tbody.innerHTML = '';
    
    // Si no hay datos
    if (agentes.length === 0 || fechas.length === 0) {
        thead.innerHTML = '<tr><th colspan="2" style="text-align: center;">No hay datos para mostrar</th></tr>';
        return;
    }
    
    // Encabezado
    let encabezadoHTML = '<tr><th>Agente</th>';
    
    fechas.forEach(fecha => {
        const fechaObj = new Date(fecha);
        const fechaFormateada = fechaObj.toLocaleDateString('es-ES', {
            day: '2-digit',
            month: '2-digit'
        });
        encabezadoHTML += `<th class="text-center">${fechaFormateada}</th>`;
    });
    
    encabezadoHTML += '<th class="text-center total-column">TOTAL</th></tr>';
    thead.innerHTML = encabezadoHTML;
    
    // Cuerpo
    tbody.innerHTML = '';
    
    // Totales por día
    const totalesDia = {};
    fechas.forEach(fecha => totalesDia[fecha] = 0);
    
    // Filas por agente
    agentes.forEach(agente => {
        const fila = document.createElement('tr');
        let filaHTML = `<td class="agent-name"><strong>${agente}</strong></td>`;
        
        let totalAgente = 0;
        
        fechas.forEach(fecha => {
            const conteo = agentesData[agente].diario[fecha] || 0;
            totalAgente += conteo;
            totalesDia[fecha] += conteo;
            filaHTML += `<td class="text-center">${conteo.toLocaleString()}</td>`;
        });
        
        filaHTML += `<td class="text-center agent-total"><strong>${totalAgente.toLocaleString()}</strong></td>`;
        fila.innerHTML = filaHTML;
        tbody.appendChild(fila);
    });
    
    // Fila de totales
    const filaTotales = document.createElement('tr');
    filaTotales.className = 'total-row';
    
    let totalesHTML = '<td><strong>TOTAL DIARIO</strong></td>';
    let totalGeneral = 0;
    
    fechas.forEach(fecha => {
        totalesHTML += `<td class="text-center day-total"><strong>${totalesDia[fecha].toLocaleString()}</strong></td>`;
        totalGeneral += totalesDia[fecha];
    });
    
    totalesHTML += `<td class="text-center"><strong>${totalGeneral.toLocaleString()}</strong></td>`;
    filaTotales.innerHTML = totalesHTML;
    tbody.appendChild(filaTotales);
}

// Exportar a Excel (versión corregida - mantiene todos los agentes)
async function exportarLlamadasDia(datos, nombreArchivo) {
    try {
        const agentesData = datos.agentesData;
        const fechas = datos.fechas;
        const agentes = Object.keys(agentesData).sort();
        
        // Crear workbook
        const workbook = new ExcelJS.Workbook();
        
        // --- HOJA 1: REPORTE LLAMADAS COMPLETO ---
        const worksheet = workbook.addWorksheet('Reporte Llamadas Completo');
        
        // Título
        worksheet.mergeCells('A1:C1');
        const titleCell = worksheet.getCell('A1');
        titleCell.value = 'REPORTE DE LLAMADAS POR DÍA (COMPLETO)';
        titleCell.style = {
            font: { bold: true, size: 16, color: { argb: 'FFFFFF' } },
            fill: { type: 'pattern', pattern: 'solid', fgColor: { argb: '2C3E50' } },
            alignment: { horizontal: 'center', vertical: 'middle' }
        };
        
        // Estadísticas
        worksheet.addRow([]);
        const statsRow = worksheet.addRow(['Estadísticas Generales - COMPLETO', '', '']);
        statsRow.font = { bold: true };
        
        worksheet.addRow(['Total Llamadas (con duplicados)', datos.todosRegistros.length, '']);
        worksheet.addRow(['Agentes Únicos', agentes.length, '']);
        
        if (fechas.length > 0) {
            const primeraFecha = new Date(fechas[0]).toLocaleDateString('es-ES');
            const ultimaFecha = new Date(fechas[fechas.length - 1]).toLocaleDateString('es-ES');
            worksheet.addRow(['Rango de Fechas', `${primeraFecha} - ${ultimaFecha}`, '']);
        }
        
        worksheet.addRow([]);
        
        // Encabezados de tabla
        const headers = ['Agente', ...fechas.map(f => {
            const fechaObj = new Date(f);
            return fechaObj.toLocaleDateString('es-ES', { day: '2-digit', month: '2-digit' });
        }), 'TOTAL'];
        
        const headerRow = worksheet.addRow(headers);
        
        headerRow.eachCell((cell) => {
            cell.style = {
                font: { bold: true, color: { argb: 'FFFFFF' } },
                fill: { type: 'pattern', pattern: 'solid', fgColor: { argb: '3498DB' } },
                alignment: { horizontal: 'center' }
            };
        });
        
        // Datos por agente (COMPLETO)
        const totalesDia = {};
        fechas.forEach(f => totalesDia[f] = 0);
        
        agentes.forEach((agente, index) => {
            const fila = [agente];
            let totalAgente = 0;
            
            fechas.forEach(fecha => {
                const conteo = agentesData[agente].diario[fecha] || 0;
                fila.push(conteo);
                totalAgente += conteo;
                totalesDia[fecha] += conteo;
            });
            
            fila.push(totalAgente);
            const dataRow = worksheet.addRow(fila);
            
            // Colores alternados
            const fillColor = index % 2 === 0 ? 'F8F9FA' : 'FFFFFF';
            dataRow.eachCell((cell) => {
                cell.style = {
                    fill: { type: 'pattern', pattern: 'solid', fgColor: { argb: fillColor } },
                    alignment: { horizontal: 'center' }
                };
            });
        });
        
        // Fila de totales (COMPLETO)
        worksheet.addRow([]);
        const totalesFila = ['TOTAL DIARIO'];
        let totalGeneral = 0;
        
        fechas.forEach(fecha => {
            totalesFila.push(totalesDia[fecha]);
            totalGeneral += totalesDia[fecha];
        });
        
        totalesFila.push(totalGeneral);
        const totalRow = worksheet.addRow(totalesFila);
        
        totalRow.eachCell((cell) => {
            cell.style = {
                font: { bold: true },
                fill: { type: 'pattern', pattern: 'solid', fgColor: { argb: 'F1C40F' } },
                alignment: { horizontal: 'center' }
            };
        });
        
        // Autoajustar columnas
        worksheet.columns.forEach((column, index) => {
            let maxLength = 0;
            column.eachCell({ includeEmpty: true }, (cell) => {
                const length = cell.value ? cell.value.toString().length : 10;
                if (length > maxLength) maxLength = length;
            });
            column.width = Math.min(maxLength + 2, 50);
        });
        
        // --- HOJA 2: DETALLE COMPLETO ---
        if (datos.todosRegistros.length > 0) {
            const detalleSheet = workbook.addWorksheet('Detalle Completo');
            
            // Obtener columnas del primer registro
            const primerRegistro = datos.todosRegistros[0].registroCompleto;
            const columnas = Object.keys(primerRegistro);
            
            // Encabezados
            const detalleHeaders = [...columnas, 'Fecha Procesada'];
            const detalleHeaderRow = detalleSheet.addRow(detalleHeaders);
            
            detalleHeaderRow.eachCell((cell) => {
                cell.style = {
                    font: { bold: true },
                    fill: { type: 'pattern', pattern: 'solid', fgColor: { argb: 'ECF0F1' } }
                };
            });
            
            // Datos completos (todos los registros)
            datos.todosRegistros.forEach(registro => {
                const fila = columnas.map(col => registro.registroCompleto[col]);
                fila.push(registro.fecha);
                detalleSheet.addRow(fila);
            });
            
            // Autoajustar columnas
            detalleSheet.columns.forEach((column) => {
                let maxLength = 0;
                column.eachCell({ includeEmpty: true }, (cell) => {
                    const length = cell.value ? cell.value.toString().length : 10;
                    if (length > maxLength) maxLength = length;
                });
                column.width = Math.min(maxLength + 2, 30);
            });
            
            // Agregar estadísticas
            detalleSheet.addRow([]);
            detalleSheet.addRow(['ESTADÍSTICAS DEL DETALLE (COMPLETO)', '', '', '', '', '', '', '']);
            detalleSheet.addRow(['Total de registros (con duplicados):', datos.todosRegistros.length]);
        }
        
        // Variables que necesitamos fuera del bloque if
        let documentosConflictivos = [];
        let registrosUnicos = [];
        let campoUnico = null;
        let columnas = [];
        
        // --- HOJA 3: REPORTE LLAMADAS SIN REPETIR ---
        // Primero necesitamos procesar los datos sin duplicados
        if (datos.todosRegistros.length > 0) {
            // Identificar campo único
            const primerRegistro = datos.todosRegistros[0].registroCompleto;
            columnas = Object.keys(primerRegistro);
            
            const posiblesCamposUnicos = ['Número Credito', 'Numero Credito', 'Número Crédito', 
                                          'Documento', 'Nro Credito', 'Credito', 'Credit', 
                                          'NroDocumento', 'NúmeroDocumento', 'Cédula',
                                          'Identificación', 'ID Cliente'];
            
            for (const campo of posiblesCamposUnicos) {
                if (columnas.includes(campo)) {
                    campoUnico = campo;
                    break;
                }
            }
            
            // 1. Filtrar registros únicos
            registrosUnicos = [];
            const valoresUnicos = new Set();
            
            // 2. Detectar documentos gestionados por múltiples agentes
            const documentosMultiplesAgentes = new Map(); // documento -> [agentes]
            
            datos.todosRegistros.forEach(registro => {
                let valorUnico = null;
                
                if (campoUnico) {
                    valorUnico = registro.registroCompleto[campoUnico];
                    
                    // Registrar qué agente gestionó este documento
                    if (valorUnico) {
                        if (!documentosMultiplesAgentes.has(valorUnico.toString())) {
                            documentosMultiplesAgentes.set(valorUnico.toString(), new Set());
                        }
                        documentosMultiplesAgentes.get(valorUnico.toString()).add(registro.agente);
                    }
                } else {
                    // Si no hay campo específico, usar combinación de varios campos
                    const camposAlternativos = ['Agente', 'Fecha', 'Cliente', 'Nombre'];
                    valorUnico = camposAlternativos
                        .filter(campo => registro.registroCompleto[campo])
                        .map(campo => registro.registroCompleto[campo])
                        .join('_');
                }
                
                if (valorUnico && !valoresUnicos.has(valorUnico.toString())) {
                    valoresUnicos.add(valorUnico.toString());
                    registrosUnicos.push(registro);
                }
            });
            
            // Filtrar documentos con múltiples agentes (2 o más)
            documentosConflictivos = [];
            documentosMultiplesAgentes.forEach((agentesSet, documento) => {
                if (agentesSet.size > 1) {
                    documentosConflictivos.push({
                        documento: documento,
                        agentes: Array.from(agentesSet),
                        cantidadAgentes: agentesSet.size
                    });
                }
            });
            
            // Ordenar por cantidad de agentes (descendente)
            documentosConflictivos.sort((a, b) => b.cantidadAgentes - a.cantidadAgentes);
            
            // 3. Crear estructura de datos para reporte sin duplicados
            // Inicializar TODOS los agentes (incluso si tienen 0 registros únicos)
            const agentesDataUnicos = {};
            const fechasSetUnicos = new Set();
            
            agentes.forEach(agente => {
                agentesDataUnicos[agente] = {
                    diario: {},
                    total: 0
                };
            });
            
            // Procesar registros únicos
            registrosUnicos.forEach(registro => {
                const agenteNombre = registro.agente;
                const fechaStr = registro.fecha;
                
                if (!agentesDataUnicos[agenteNombre].diario[fechaStr]) {
                    agentesDataUnicos[agenteNombre].diario[fechaStr] = 0;
                }
                
                agentesDataUnicos[agenteNombre].diario[fechaStr]++;
                agentesDataUnicos[agenteNombre].total++;
                fechasSetUnicos.add(fechaStr);
            });
            
            const fechasUnicas = Array.from(fechasSetUnicos).sort();
            // Mantener MISMO orden de agentes que en el reporte completo
            const agentesUnicos = agentes; // Usar misma lista
            
            // Crear hoja de reporte sin duplicados
            const reporteUnicoSheet = workbook.addWorksheet('Reporte Sin Duplicados');
            
            // Título
            reporteUnicoSheet.mergeCells('A1:D1');
            const titleUnicoCell = reporteUnicoSheet.getCell('A1');
            titleUnicoCell.value = 'REPORTE DE LLAMADAS POR DÍA (SIN DUPLICADOS)';
            titleUnicoCell.style = {
                font: { bold: true, size: 16, color: { argb: 'FFFFFF' } },
                fill: { type: 'pattern', pattern: 'solid', fgColor: { argb: '27AE60' } },
                alignment: { horizontal: 'center', vertical: 'middle' }
            };
            
            // Estadísticas específicas para sin duplicados
            reporteUnicoSheet.addRow([]);
            const statsUnicoRow = reporteUnicoSheet.addRow(['Estadísticas - SIN DUPLICADOS', '', '', '']);
            statsUnicoRow.font = { bold: true };
            
            reporteUnicoSheet.addRow(['Total registros únicos:', registrosUnicos.length, '', '']);
            reporteUnicoSheet.addRow(['Total registros originales:', datos.todosRegistros.length, '', '']);
            reporteUnicoSheet.addRow(['Registros eliminados (duplicados):', datos.todosRegistros.length - registrosUnicos.length, '', '']);
            reporteUnicoSheet.addRow(['Porcentaje de reducción:', `${((1 - (registrosUnicos.length / datos.todosRegistros.length)) * 100).toFixed(2)}%`, '', '']);
            reporteUnicoSheet.addRow(['Campo utilizado:', campoUnico || 'Combinación de campos', '', '']);
            reporteUnicoSheet.addRow(['Total agentes:', agentesUnicos.length, '', '']);
            reporteUnicoSheet.addRow(['Agentes con 0 registros únicos:', 
                agentesUnicos.filter(a => agentesDataUnicos[a].total === 0).length, '', '']);
            
            // Información de documentos gestionados por múltiples agentes
            if (documentosConflictivos.length > 0) {
                reporteUnicoSheet.addRow([]);
                reporteUnicoSheet.addRow(['DOCUMENTOS GESTIONADOS POR MÚLTIPLES AGENTES', '', '', '']);
                reporteUnicoSheet.addRow(['Total documentos con múltiples agentes:', documentosConflictivos.length, '', '']);
                reporteUnicoSheet.addRow(['Documento con más agentes:', 
                    documentosConflictivos.length > 0 ? 
                    `${documentosConflictivos[0].documento} (${documentosConflictivos[0].cantidadAgentes} agentes)` : 
                    'N/A', '', '']);
            }
            
            reporteUnicoSheet.addRow([]);
            
            // Encabezados de tabla sin duplicados
            if (fechasUnicas.length > 0) {
                const headersUnicos = ['Agente', ...fechasUnicas.map(f => {
                    const fechaObj = new Date(f);
                    return fechaObj.toLocaleDateString('es-ES', { day: '2-digit', month: '2-digit' });
                }), 'TOTAL'];
                
                const headerUnicoRow = reporteUnicoSheet.addRow(headersUnicos);
                
                headerUnicoRow.eachCell((cell) => {
                    cell.style = {
                        font: { bold: true, color: { argb: 'FFFFFF' } },
                        fill: { type: 'pattern', pattern: 'solid', fgColor: { argb: '27AE60' } },
                        alignment: { horizontal: 'center' }
                    };
                });
                
                // Datos por agente (SIN DUPLICADOS) - TODOS los agentes
                const totalesDiaUnicos = {};
                fechasUnicas.forEach(f => totalesDiaUnicos[f] = 0);
                
                agentesUnicos.forEach((agente, index) => {
                    const fila = [agente];
                    let totalAgente = 0;
                    
                    fechasUnicas.forEach(fecha => {
                        const conteo = agentesDataUnicos[agente].diario[fecha] || 0;
                        fila.push(conteo);
                        totalAgente += conteo;
                        totalesDiaUnicos[fecha] += conteo;
                    });
                    
                    fila.push(totalAgente);
                    const dataRow = reporteUnicoSheet.addRow(fila);
                    
                    // Resaltar agentes con 0 registros únicos
                    if (totalAgente === 0) {
                        dataRow.eachCell((cell) => {
                            cell.style = {
                                fill: { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFEBEE' } },
                                alignment: { horizontal: 'center' },
                                font: { italic: true, color: { argb: '757575' } }
                            };
                        });
                    } else {
                        // Colores alternados normales
                        const fillColor = index % 2 === 0 ? 'F0F8F0' : 'FFFFFF';
                        dataRow.eachCell((cell) => {
                            cell.style = {
                                fill: { type: 'pattern', pattern: 'solid', fgColor: { argb: fillColor } },
                                alignment: { horizontal: 'center' }
                            };
                        });
                    }
                });
                
                // Fila de totales (SIN DUPLICADOS)
                reporteUnicoSheet.addRow([]);
                const totalesFilaUnicos = ['TOTAL DIARIO'];
                let totalGeneralUnicos = 0;
                
                fechasUnicas.forEach(fecha => {
                    totalesFilaUnicos.push(totalesDiaUnicos[fecha]);
                    totalGeneralUnicos += totalesDiaUnicos[fecha];
                });
                
                totalesFilaUnicos.push(totalGeneralUnicos);
                const totalUnicoRow = reporteUnicoSheet.addRow(totalesFilaUnicos);
                
                totalUnicoRow.eachCell((cell) => {
                    cell.style = {
                        font: { bold: true },
                        fill: { type: 'pattern', pattern: 'solid', fgColor: { argb: 'F39C12' } },
                        alignment: { horizontal: 'center' }
                    };
                });
                
                // Autoajustar columnas
                reporteUnicoSheet.columns.forEach((column, index) => {
                    let maxLength = 0;
                    column.eachCell({ includeEmpty: true }, (cell) => {
                        const length = cell.value ? cell.value.toString().length : 10;
                        if (length > maxLength) maxLength = length;
                    });
                    column.width = Math.min(maxLength + 2, 50);
                });
            }
            
            // --- HOJA 4: DETALLE SIN REPETIR ---
            const detalleUnicoSheet = workbook.addWorksheet('Detalle Sin Repetir');
            
            // Encabezados
            const detalleUnicoHeaders = [...columnas, 'Fecha Procesada'];
            const detalleUnicoHeaderRow = detalleUnicoSheet.addRow(detalleUnicoHeaders);
            
            detalleUnicoHeaderRow.eachCell((cell) => {
                cell.style = {
                    font: { bold: true },
                    fill: { type: 'pattern', pattern: 'solid', fgColor: { argb: 'D5F4E6' } }
                };
            });
            
            // Ordenar registros únicos por fecha y agente
            registrosUnicos.sort((a, b) => {
                if (a.fecha === b.fecha) {
                    return a.agente.localeCompare(b.agente);
                }
                return a.fecha.localeCompare(b.fecha);
            });
            
            // Agregar datos únicos a la hoja
            registrosUnicos.forEach(registro => {
                const fila = columnas.map(col => registro.registroCompleto[col]);
                fila.push(registro.fecha);
                detalleUnicoSheet.addRow(fila);
            });
            
            // Autoajustar columnas
            detalleUnicoSheet.columns.forEach((column) => {
                let maxLength = 0;
                column.eachCell({ includeEmpty: true }, (cell) => {
                    const length = cell.value ? cell.value.toString().length : 10;
                    if (length > maxLength) maxLength = length;
                });
                column.width = Math.min(maxLength + 2, 30);
            });
            
            // Agregar estadísticas
            detalleUnicoSheet.addRow([]);
            detalleUnicoSheet.addRow(['ESTADÍSTICAS - DETALLE SIN DUPLICADOS', '', '', '', '', '', '', '']);
            detalleUnicoSheet.addRow(['Total registros únicos:', registrosUnicos.length]);
            detalleUnicoSheet.addRow(['Campo utilizado para eliminar duplicados:', campoUnico || 'Combinación de campos']);
            detalleUnicoSheet.addRow(['Fecha de generación:', new Date().toLocaleString()]);
        }
        
        // --- HOJA 5: DOCUMENTOS CON MÚLTIPLES AGENTES (OPCIONAL) ---
        // SOLO si hay documentos gestionados por múltiples agentes
        if (documentosConflictivos && documentosConflictivos.length > 0) {
            const conflictosSheet = workbook.addWorksheet('Doc Multiples Agentes');
            
            // Título
            conflictosSheet.mergeCells('A1:D1');
            const titleConflictCell = conflictosSheet.getCell('A1');
            titleConflictCell.value = 'DOCUMENTOS GESTIONADOS POR MÚLTIPLES AGENTES';
            titleConflictCell.style = {
                font: { bold: true, size: 16, color: { argb: 'FFFFFF' } },
                fill: { type: 'pattern', pattern: 'solid', fgColor: { argb: 'E74C3C' } },
                alignment: { horizontal: 'center', vertical: 'middle' }
            };
            
            conflictosSheet.addRow([]);
            conflictosSheet.addRow(['Estadísticas de Conflictos', '', '', '']);
            conflictosSheet.addRow(['Total documentos con múltiples agentes:', documentosConflictivos.length]);
            conflictosSheet.addRow(['Promedio de agentes por documento:', 
                (documentosConflictivos.reduce((sum, doc) => sum + doc.cantidadAgentes, 0) / documentosConflictivos.length).toFixed(2)]);
            
            conflictosSheet.addRow([]);
            
            // Encabezados
            const conflictHeaders = ['Documento/Crédito', 'Cantidad de Agentes', 'Agentes Involucrados'];
            const conflictHeaderRow = conflictosSheet.addRow(conflictHeaders);
            
            conflictHeaderRow.eachCell((cell) => {
                cell.style = {
                    font: { bold: true, color: { argb: 'FFFFFF' } },
                    fill: { type: 'pattern', pattern: 'solid', fgColor: { argb: 'E74C3C' } }
                };
            });
            
            // Datos de conflictos
            documentosConflictivos.forEach((doc, index) => {
                const fila = [
                    doc.documento,
                    doc.cantidadAgentes,
                    doc.agentes.join(', ')
                ];
                
                const dataRow = conflictosSheet.addRow(fila);
                
                // Resaltar según cantidad de agentes
                let bgColor = 'FFFFFF';
                if (doc.cantidadAgentes >= 4) {
                    bgColor = 'FFEBEE'; // Rojo claro para muchos conflictos
                } else if (doc.cantidadAgentes === 3) {
                    bgColor = 'FFF3E0'; // Naranja claro
                } else if (doc.cantidadAgentes === 2) {
                    bgColor = 'F3E5F5'; // Púrpura claro
                }
                
                dataRow.eachCell((cell) => {
                    cell.style = {
                        fill: { type: 'pattern', pattern: 'solid', fgColor: { argb: bgColor } }
                    };
                });
            });
            
            // Autoajustar columnas
            conflictosSheet.columns.forEach((column) => {
                let maxLength = 0;
                column.eachCell({ includeEmpty: true }, (cell) => {
                    const length = cell.value ? cell.value.toString().length : 10;
                    if (length > maxLength) maxLength = length;
                });
                column.width = Math.min(maxLength + 2, 40);
            });
        }
        
        // Descargar
        const buffer = await workbook.xlsx.writeBuffer();
        const blob = new Blob([buffer], {
            type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
        });
        
        const url = window.URL.createObjectURL(blob);
        const a = document.createElement('a');
        a.href = url;
        
        const nombreBase = nombreArchivo.replace(/\.[^/.]+$/, "");
        const fechaHoy = new Date().toISOString().split('T')[0];
        const hojasCount = documentosConflictivos && documentosConflictivos.length > 0 ? 5 : 4;
        a.download = `${nombreBase}_reporte_${hojasCount}hojas_${fechaHoy}.xlsx`;
        
        document.body.appendChild(a);
        a.click();
        document.body.removeChild(a);
        window.URL.revokeObjectURL(url);
        
        let mensaje = 'Reporte exportado exitosamente con 4 hojas';
        if (documentosConflictivos && documentosConflictivos.length > 0) {
            mensaje = `Reporte exportado con ${hojasCount} hojas (incluye ${documentosConflictivos.length} documentos con múltiples agentes)`;
        }
        
        showStatus(mensaje, 'success');
        
    } catch (error) {
        console.error('Error al exportar reporte de llamadas:', error);
        showStatus('Error al exportar el reporte', 'error');
    }
}

// Resetear reporte
function resetearReporteLlamadasCompleto() {
    llamadasDiaContainer.style.display = 'none';
    results.style.display = 'none';
    resultsContent.innerHTML = '';
    processedFiles = [];
    fileInput.value = '';
    
    // Resetear estadísticas
    document.getElementById('totalCalls').textContent = '0';
    document.getElementById('uniqueAgents').textContent = '0';
    document.getElementById('dateRange').textContent = '-';
    document.getElementById('topAgents').innerHTML = '';
    
    // Limpiar tabla
    document.getElementById('llamadasTableHeader').innerHTML = '';
    document.getElementById('llamadasTableBody').innerHTML = '';
    
    showStatus('Reporte reiniciado. Puede seleccionar otro archivo.', 'info');
}

// Exportar al scope global
window.procesarLlamadasDia = procesarLlamadasDia;
window.mostrarReporteLlamadasCompleto = mostrarReporteLlamadasCompleto;
window.exportarLlamadasDia = exportarLlamadasDia;
window.resetearReporteLlamadasCompleto = resetearReporteLlamadasCompleto;
window.resetearReporteLlamadas = resetearReporteLlamadasCompleto; // Alias para compatibilidad
