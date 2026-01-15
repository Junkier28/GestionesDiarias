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

// Exportar a Excel (versión con resumen por nivel 1)
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

        // --- HOJA 3: REPORTE LLAMADAS SIN REPETIR (POR DÍA) ---
        // Primero necesitamos procesar los datos sin duplicados POR DÍA
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

            // 1. Filtrar registros únicos POR DÍA
            registrosUnicos = [];
            const valoresUnicosPorDia = new Map(); // fecha -> Set de valores únicos

            // 2. Detectar documentos gestionados por múltiples agentes POR DÍA
            const documentosMultiplesAgentesPorDia = new Map(); // fecha_documento -> [agentes]

            datos.todosRegistros.forEach(registro => {
                const fechaStr = registro.fecha;
                let valorUnico = null;

                if (campoUnico) {
                    valorUnico = registro.registroCompleto[campoUnico];

                    if (valorUnico) {
                        // Crear clave única por fecha + documento
                        const clave = `${fechaStr}_${valorUnico.toString()}`;

                        // Registrar qué agente gestionó este documento en esta fecha
                        if (!documentosMultiplesAgentesPorDia.has(clave)) {
                            documentosMultiplesAgentesPorDia.set(clave, {
                                fecha: fechaStr,
                                documento: valorUnico.toString(),
                                agentes: new Set()
                            });
                        }
                        documentosMultiplesAgentesPorDia.get(clave).agentes.add(registro.agente);
                    }
                }

                // Inicializar Set para esta fecha si no existe
                if (!valoresUnicosPorDia.has(fechaStr)) {
                    valoresUnicosPorDia.set(fechaStr, new Set());
                }

                // Si no existe este valor en esta fecha, agregarlo
                if (valorUnico && !valoresUnicosPorDia.get(fechaStr).has(valorUnico.toString())) {
                    valoresUnicosPorDia.get(fechaStr).add(valorUnico.toString());
                    registrosUnicos.push(registro);
                }
            });

            // Filtrar documentos con múltiples agentes POR DÍA
            documentosConflictivos = [];
            documentosMultiplesAgentesPorDia.forEach((info, clave) => {
                if (info.agentes.size > 1) {
                    documentosConflictivos.push({
                        fecha: info.fecha,
                        documento: info.documento,
                        agentes: Array.from(info.agentes),
                        cantidadAgentes: info.agentes.size
                    });
                }
            });

            // Ordenar por fecha y cantidad de agentes
            documentosConflictivos.sort((a, b) => {
                if (a.fecha === b.fecha) {
                    return b.cantidadAgentes - a.cantidadAgentes;
                }
                return a.fecha.localeCompare(b.fecha);
            });

            // 3. Crear estructura de datos para reporte sin duplicados POR DÍA
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
            titleUnicoCell.value = 'REPORTE DE LLAMADAS POR DÍA (SIN DUPLICADOS - POR DÍA)';
            titleUnicoCell.style = {
                font: { bold: true, size: 16, color: { argb: 'FFFFFF' } },
                fill: { type: 'pattern', pattern: 'solid', fgColor: { argb: '27AE60' } },
                alignment: { horizontal: 'center', vertical: 'middle' }
            };

            // Nota explicativa
            reporteUnicoSheet.addRow([]);
            reporteUnicoSheet.addRow(['NOTA: Los duplicados se eliminan por día. Si un documento aparece múltiples veces en el mismo día, cuenta como 1.', '', '', '']);
            reporteUnicoSheet.addRow(['Si aparece en diferentes días, cuenta como 1 en cada día.', '', '', '']);

            // Estadísticas específicas para sin duplicados
            reporteUnicoSheet.addRow([]);
            const statsUnicoRow = reporteUnicoSheet.addRow(['Estadísticas - SIN DUPLICADOS', '', '', '']);
            statsUnicoRow.font = { bold: true };

            reporteUnicoSheet.addRow(['Total registros únicos:', registrosUnicos.length, '', '']);
            reporteUnicoSheet.addRow(['Total registros originales:', datos.todosRegistros.length, '', '']);
            reporteUnicoSheet.addRow(['Registros eliminados (duplicados por día):', datos.todosRegistros.length - registrosUnicos.length, '', '']);
            reporteUnicoSheet.addRow(['Porcentaje de reducción:', `${((1 - (registrosUnicos.length / datos.todosRegistros.length)) * 100).toFixed(2)}%`, '', '']);
            reporteUnicoSheet.addRow(['Campo utilizado:', campoUnico || 'Combinación de campos', '', '']);
            reporteUnicoSheet.addRow(['Total agentes:', agentesUnicos.length, '', '']);
            reporteUnicoSheet.addRow(['Agentes con 0 registros únicos:',
                agentesUnicos.filter(a => agentesDataUnicos[a].total === 0).length, '', '']);

            // Información de documentos gestionados por múltiples agentes POR DÍA
            if (documentosConflictivos.length > 0) {
                reporteUnicoSheet.addRow([]);
                reporteUnicoSheet.addRow(['DOCUMENTOS GESTIONADOS POR MÚLTIPLES AGENTES (POR DÍA)', '', '', '']);
                reporteUnicoSheet.addRow(['Total documentos con múltiples agentes en un mismo día:', documentosConflictivos.length, '', '']);

                // Agrupar por fecha para mostrar mejor
                const conflictosPorFecha = {};
                documentosConflictivos.forEach(conflicto => {
                    if (!conflictosPorFecha[conflicto.fecha]) {
                        conflictosPorFecha[conflicto.fecha] = 0;
                    }
                    conflictosPorFecha[conflicto.fecha]++;
                });

                // Mostrar los días con más conflictos
                const diasConConflictos = Object.entries(conflictosPorFecha)
                    .sort((a, b) => b[1] - a[1])
                    .slice(0, 3);

                if (diasConConflictos.length > 0) {
                    diasConConflictos.forEach(([fecha, cantidad], index) => {
                        const fechaFormateada = new Date(fecha).toLocaleDateString('es-ES');
                        reporteUnicoSheet.addRow([`${index === 0 ? 'Día con más conflictos:' : 'Siguiente:'}`, `${fechaFormateada} (${cantidad} documentos)`, '', '']);
                    });
                }
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

            // Título con explicación
            detalleUnicoSheet.mergeCells('A1:D1');
            const titleDetalleCell = detalleUnicoSheet.getCell('A1');
            titleDetalleCell.value = 'DETALLE SIN DUPLICADOS (POR DÍA)';
            titleDetalleCell.style = {
                font: { bold: true, size: 16, color: { argb: 'FFFFFF' } },
                fill: { type: 'pattern', pattern: 'solid', fgColor: { argb: 'D5F4E6' } },
                alignment: { horizontal: 'center', vertical: 'middle' }
            };

            detalleUnicoSheet.addRow([]);
            detalleUnicoSheet.addRow(['NOTA: Se muestran solo los registros únicos por día. Si un documento aparece múltiples veces en el mismo día,', '', '', '']);
            detalleUnicoSheet.addRow(['solo se incluye el primer registro de ese documento para ese día.', '', '', '']);

            detalleUnicoSheet.addRow([]);

            // Encabezados
            const detalleUnicoHeaders = [...columnas, 'Fecha Procesada'];
            const detalleUnicoHeaderRow = detalleUnicoSheet.addRow(detalleUnicoHeaders);

            detalleUnicoHeaderRow.eachCell((cell) => {
                cell.style = {
                    font: { bold: true },
                    fill: { type: 'pattern', pattern: 'solid', fgColor: { argb: 'A8E6CF' } }
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
            detalleUnicoSheet.addRow(['ESTADÍSTICAS - DETALLE SIN DUPLICADOS (POR DÍA)', '', '', '', '', '', '', '']);
            detalleUnicoSheet.addRow(['Total registros únicos por día:', registrosUnicos.length]);
            detalleUnicoSheet.addRow(['Campo utilizado para eliminar duplicados:', campoUnico || 'Combinación de campos']);
            detalleUnicoSheet.addRow(['Fecha de generación:', new Date().toLocaleString()]);
        }

        // --- HOJA 5: DOCUMENTOS CON MÚLTIPLES AGENTES POR DÍA (OPCIONAL) ---
        if (documentosConflictivos && documentosConflictivos.length > 0) {
            const conflictosSheet = workbook.addWorksheet('Doc Multiples Agentes x Dia');

            // Título
            conflictosSheet.mergeCells('A1:E1');
            const titleConflictCell = conflictosSheet.getCell('A1');
            titleConflictCell.value = 'DOCUMENTOS GESTIONADOS POR MÚLTIPLES AGENTES (POR DÍA)';
            titleConflictCell.style = {
                font: { bold: true, size: 16, color: { argb: 'FFFFFF' } },
                fill: { type: 'pattern', pattern: 'solid', fgColor: { argb: 'E74C3C' } },
                alignment: { horizontal: 'center', vertical: 'middle' }
            };

            conflictosSheet.addRow([]);
            conflictosSheet.addRow(['NOTA: Documentos que fueron gestionados por 2 o más agentes en el MISMO DÍA', '', '', '', '']);

            conflictosSheet.addRow([]);
            conflictosSheet.addRow(['Estadísticas de Conflictos por Día', '', '', '', '']);

            // Estadísticas por fecha
            const conflictosPorFecha = {};
            documentosConflictivos.forEach(conflicto => {
                if (!conflictosPorFecha[conflicto.fecha]) {
                    conflictosPorFecha[conflicto.fecha] = {
                        cantidad: 0,
                        documentos: []
                    };
                }
                conflictosPorFecha[conflicto.fecha].cantidad++;
                conflictosPorFecha[conflicto.fecha].documentos.push(conflicto.documento);
            });

            // Ordenar fechas por cantidad de conflictos
            const fechasConflictos = Object.entries(conflictosPorFecha)
                .sort((a, b) => b[1].cantidad - a[1].cantidad);

            conflictosSheet.addRow(['Total días con conflictos:', fechasConflictos.length, '', '', '']);
            conflictosSheet.addRow(['Total documentos conflictivos:', documentosConflictivos.length, '', '', '']);

            if (fechasConflictos.length > 0) {
                const fechaMaxConflictos = fechasConflictos[0];
                const fechaFormateada = new Date(fechaMaxConflictos[0]).toLocaleDateString('es-ES');
                conflictosSheet.addRow(['Día con más conflictos:', `${fechaFormateada} (${fechaMaxConflictos[1].cantidad} documentos)`, '', '', '']);
            }

            conflictosSheet.addRow([]);

            // Encabezados
            const conflictHeaders = ['Fecha', 'Documento/Crédito', 'Cantidad de Agentes', 'Agentes Involucrados'];
            const conflictHeaderRow = conflictosSheet.addRow(conflictHeaders);

            conflictHeaderRow.eachCell((cell) => {
                cell.style = {
                    font: { bold: true, color: { argb: 'FFFFFF' } },
                    fill: { type: 'pattern', pattern: 'solid', fgColor: { argb: 'E74C3C' } }
                };
            });

            // Datos de conflictos
            documentosConflictivos.forEach((doc, index) => {
                const fechaFormateada = new Date(doc.fecha).toLocaleDateString('es-ES');
                const fila = [
                    fechaFormateada,
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

        // --- HOJA 6: RESUMEN POR GESTOR NIVEL 1 (UNIFICADO) ---
        if (datos.todosRegistros.length > 0) {
            // Verificar si la hoja ya existe
            let nombreHojaResumen = 'Resumen Nivel 1';
            let contador = 1;

            while (workbook.getWorksheet(nombreHojaResumen)) {
                nombreHojaResumen = `Resumen Nivel 1 (${contador})`;
                contador++;
            }

            const resumenNivel1Sheet = workbook.addWorksheet(nombreHojaResumen);

            // Título
            resumenNivel1Sheet.mergeCells('A1:I1');
            const titleNivel1Cell = resumenNivel1Sheet.getCell('A1');
            titleNivel1Cell.value = 'RESUMEN POR GESTOR - NIVEL 1';
            titleNivel1Cell.style = {
                font: { bold: true, size: 16, color: { argb: 'FFFFFF' } },
                fill: { type: 'pattern', pattern: 'solid', fgColor: { argb: '8E44AD' } },
                alignment: { horizontal: 'center', vertical: 'middle' }
            };

            // Detectar columna "Nivel 1"
            const primerRegistro = datos.todosRegistros[0].registroCompleto;
            const columnas = Object.keys(primerRegistro);

            let columnaNivel1 = null;
            for (const columna of columnas) {
                if (columna.toString().trim().toUpperCase() === 'NIVEL 1') {
                    columnaNivel1 = columna;
                    break;
                }
            }

            // Si no encuentra exacto, buscar variaciones
            if (!columnaNivel1) {
                for (const columna of columnas) {
                    const columnaUpper = columna.toString().trim().toUpperCase();
                    if (columnaUpper.includes('NIVEL') || columnaUpper.includes('LEVEL')) {
                        columnaNivel1 = columna;
                        break;
                    }
                }
            }

            // Si aún no encuentra, buscar cualquier columna que pueda tener los valores
            if (!columnaNivel1) {
                const posiblesValores = ['TELEFON', 'WHATSAPP', 'CORREO', 'TERRENO', 'CONTACTO'];
                for (const columna of columnas) {
                    // Revisar algunos registros para ver si esta columna tiene los valores
                    let tieneValores = false;
                    for (let i = 0; i < Math.min(5, datos.todosRegistros.length); i++) {
                        const valor = datos.todosRegistros[i].registroCompleto[columna];
                        if (valor) {
                            const valorStr = valor.toString().toUpperCase();
                            if (posiblesValores.some(v => valorStr.includes(v))) {
                                tieneValores = true;
                                break;
                            }
                        }
                    }
                    if (tieneValores) {
                        columnaNivel1 = columna;
                        break;
                    }
                }
            }

            // Definir TODAS las categorías de Nivel 1 unificadas
            const todasCategoriasNivel1 = [
                'GESTION TELEFONICA',
                'GESTION WHATSAPP',
                'SIN CONTACTO',
                'GESTION CORREO',
                'CONTACTO INDIRECTO',
                'CONTACTO DIRECTO',
                'GESTION TERRENO',
                'CONTACTO NO EFECTIVO', 
                'CONTACTO EFECTIVO'
            ];

            // Mapeo flexible para diferentes nombres
            const mapeoCategorias = {
                'TELEFON': 'GESTION TELEFONICA',
                'TELEPHONE': 'GESTION TELEFONICA',
                'LLAMAD': 'GESTION TELEFONICA',
                'CALL': 'GESTION TELEFONICA',
                'TELEFÓNICA': 'GESTION TELEFONICA',
                'TELEFÓNICO': 'GESTION TELEFONICA',

                'WHATSAPP': 'GESTION WHATSAPP',
                'WA': 'GESTION WHATSAPP',
                'MESSAGE': 'GESTION WHATSAPP',
                'MENSAJE': 'GESTION WHATSAPP',

                'CORREO': 'GESTION CORREO',
                'EMAIL': 'GESTION CORREO',
                'MAIL': 'GESTION CORREO',
                'E-MAIL': 'GESTION CORREO',

                'TERRENO': 'GESTION TERRENO',
                'VISITA': 'GESTION TERRENO',
                'FIELD': 'GESTION TERRENO',
                'CAMPO': 'GESTION TERRENO',

                'DIRECTO': 'CONTACTO DIRECTO',
                'PERSONA': 'CONTACTO DIRECTO',
                'DIRECT': 'CONTACTO DIRECTO',
                'PERSONAL': 'CONTACTO DIRECTO',

                'INDIRECTO': 'CONTACTO INDIRECTO',
                'INDIRECT': 'CONTACTO INDIRECTO',

                'SIN CONTACTO': 'SIN CONTACTO',
                'NO CONTACTO': 'SIN CONTACTO',
                'SIN CONTACT': 'SIN CONTACTO',
                'NO CONTACT': 'SIN CONTACTO',
                'SIN CONTAC': 'SIN CONTACTO'
            };

            // Información de la columna detectada
            resumenNivel1Sheet.addRow([]);
            resumenNivel1Sheet.addRow(['INFORMACIÓN DEL REPORTE', '', '', '', '', '', '', '', '']);
            resumenNivel1Sheet.addRow(['Columna detectada para Nivel 1:', columnaNivel1 || 'No encontrada', '', '', '', '', '', '', '']);
            resumenNivel1Sheet.addRow(['Categorías Nivel 1 (todas):', todasCategoriasNivel1.join(', '), '', '', '', '', '', '', '']);

            if (columnaNivel1) {
                // Recolectar todos los valores únicos de la columna Nivel 1 para debug
                const valoresUnicos = new Set();
                datos.todosRegistros.forEach(registro => {
                    const valor = registro.registroCompleto[columnaNivel1];
                    if (valor) {
                        valoresUnicos.add(valor.toString().trim().toUpperCase());
                    }
                });

                console.log('Valores únicos encontrados en Nivel 1:', Array.from(valoresUnicos));

                // Crear estructura para el resumen por gestor
                const resumenPorGestor = {};

                // Inicializar estructura para cada gestor
                agentes.forEach(agente => {
                    resumenPorGestor[agente] = {
                        porFecha: {},
                        totales: {}
                    };

                    // Inicializar totales por categoría
                    todasCategoriasNivel1.forEach(categoria => {
                        resumenPorGestor[agente].totales[categoria] = 0;
                    });
                    resumenPorGestor[agente].totales['TOTAL'] = 0;
                });

                // Procesar todos los registros
                let registrosProcesados = 0;
                let registrosNoClasificados = 0;

                datos.todosRegistros.forEach(registro => {
                    const agente = registro.agente;
                    const fecha = registro.fecha;
                    const registroCompleto = registro.registroCompleto;

                    // Obtener valor de Nivel 1
                    if (columnaNivel1 && registroCompleto[columnaNivel1]) {
                        const valorOriginal = registroCompleto[columnaNivel1].toString().trim();
                        const valorNivel1 = valorOriginal.toUpperCase();

                        // Encontrar la categoría correspondiente
                        let categoriaEncontrada = null;

                        // 1. Buscar coincidencia exacta con categorías principales
                        for (const categoria of todasCategoriasNivel1) {
                            if (valorNivel1 === categoria.toUpperCase()) {
                                categoriaEncontrada = categoria;
                                break;
                            }
                        }

                        // 2. Buscar por mapeo de palabras clave
                        if (!categoriaEncontrada) {
                            for (const [palabraClave, categoria] of Object.entries(mapeoCategorias)) {
                                if (valorNivel1.includes(palabraClave.toUpperCase())) {
                                    categoriaEncontrada = categoria;
                                    break;
                                }
                            }
                        }

                        // 3. Si aún no se encontró, buscar coincidencias parciales
                        if (!categoriaEncontrada) {
                            for (const categoria of todasCategoriasNivel1) {
                                const catUpper = categoria.toUpperCase();
                                if (valorNivel1.includes(catUpper) || catUpper.includes(valorNivel1)) {
                                    categoriaEncontrada = categoria;
                                    break;
                                }
                            }
                        }

                        // Si se encontró categoría, procesar
                        if (categoriaEncontrada) {
                            // Inicializar estructura para esta fecha si no existe
                            if (!resumenPorGestor[agente].porFecha[fecha]) {
                                resumenPorGestor[agente].porFecha[fecha] = {};
                                todasCategoriasNivel1.forEach(categoria => {
                                    resumenPorGestor[agente].porFecha[fecha][categoria] = 0;
                                });
                                resumenPorGestor[agente].porFecha[fecha]['TOTAL'] = 0;
                            }

                            // Incrementar contadores
                            resumenPorGestor[agente].porFecha[fecha][categoriaEncontrada]++;
                            resumenPorGestor[agente].porFecha[fecha]['TOTAL']++;
                            resumenPorGestor[agente].totales[categoriaEncontrada]++;
                            resumenPorGestor[agente].totales['TOTAL']++;

                            registrosProcesados++;
                        } else {
                            registrosNoClasificados++;
                            console.log('Valor no clasificado:', valorOriginal);
                        }
                    }
                });

                console.log(`Registros procesados: ${registrosProcesados}, No clasificados: ${registrosNoClasificados}`);

                // Determinar qué categorías realmente tienen datos
                const categoriasConDatos = todasCategoriasNivel1.filter(categoria => {
                    return agentes.some(agente => resumenPorGestor[agente].totales[categoria] > 0);
                });

                // Crear tabla resumen por gestor
                resumenNivel1Sheet.addRow([]);
                resumenNivel1Sheet.addRow(['RESUMEN POR GESTOR - TOTALES', '', '', '', '', '', '', '', '']);

                // Usar solo categorías que tienen datos
                const categoriasAMostrar = categoriasConDatos.length > 0 ? categoriasConDatos : todasCategoriasNivel1;

                // Encabezados de la tabla
                const encabezados = ['Agente', ...categoriasAMostrar, 'TOTAL'];
                const headerNivel1Row = resumenNivel1Sheet.addRow(encabezados);

                headerNivel1Row.eachCell((cell) => {
                    cell.style = {
                        font: { bold: true, color: { argb: 'FFFFFF' } },
                        fill: { type: 'pattern', pattern: 'solid', fgColor: { argb: '9B59B6' } },
                        alignment: { horizontal: 'center' }
                    };
                });

                // Datos por agente
                agentes.forEach((agente, index) => {
                    const fila = [agente];

                    // Agregar totales por categoría
                    categoriasAMostrar.forEach(categoria => {
                        fila.push(resumenPorGestor[agente].totales[categoria] || 0);
                    });

                    // Agregar total general
                    fila.push(resumenPorGestor[agente].totales['TOTAL'] || 0);

                    const dataRow = resumenNivel1Sheet.addRow(fila);

                    // Colores alternados
                    const fillColor = index % 2 === 0 ? 'F5EEF8' : 'FFFFFF';
                    dataRow.eachCell((cell) => {
                        cell.style = {
                            fill: { type: 'pattern', pattern: 'solid', fgColor: { argb: fillColor } },
                            alignment: { horizontal: 'center' }
                        };
                    });
                });

                // Fila de totales generales
                resumenNivel1Sheet.addRow([]);
                const totalesGeneralesFila = ['TOTAL GENERAL'];

                categoriasAMostrar.forEach(categoria => {
                    const totalCategoria = agentes.reduce((sum, agente) =>
                        sum + (resumenPorGestor[agente].totales[categoria] || 0), 0);
                    totalesGeneralesFila.push(totalCategoria);
                });

                const totalGeneral = agentes.reduce((sum, agente) =>
                    sum + (resumenPorGestor[agente].totales['TOTAL'] || 0), 0);
                totalesGeneralesFila.push(totalGeneral);

                const totalGeneralesRow = resumenNivel1Sheet.addRow(totalesGeneralesFila);

                totalGeneralesRow.eachCell((cell) => {
                    cell.style = {
                        font: { bold: true },
                        fill: { type: 'pattern', pattern: 'solid', fgColor: { argb: 'D2B4DE' } },
                        alignment: { horizontal: 'center' }
                    };
                });

                // Crear tablas detalladas por fecha para cada agente
                resumenNivel1Sheet.addRow([]);
                resumenNivel1Sheet.addRow(['DETALLE POR FECHA Y GESTOR', '', '', '', '', '', '', '', '']);
                resumenNivel1Sheet.addRow(['NOTA: Se muestra el desglose de gestiones por día para cada gestor', '', '', '', '', '', '', '', '']);

                agentes.forEach(agente => {
                    if (resumenPorGestor[agente].totales['TOTAL'] > 0) {
                        resumenNivel1Sheet.addRow([]);
                        resumenNivel1Sheet.addRow([`GESTOR: ${agente}`, '', '', '', '', '', '', '', '']);

                        // Encabezados para el detalle por fecha
                        const detalleHeaders = ['Fecha', ...categoriasAMostrar, 'TOTAL DIA'];
                        const detalleHeaderRow = resumenNivel1Sheet.addRow(detalleHeaders);

                        detalleHeaderRow.eachCell((cell) => {
                            cell.style = {
                                font: { bold: true },
                                fill: { type: 'pattern', pattern: 'solid', fgColor: { argb: 'E8DAEF' } }
                            };
                        });

                        // Ordenar fechas para este agente
                        const fechasAgente = Object.keys(resumenPorGestor[agente].porFecha || {}).sort();

                        // Agregar datos por fecha
                        fechasAgente.forEach(fecha => {
                            if (resumenPorGestor[agente].porFecha[fecha] && resumenPorGestor[agente].porFecha[fecha]['TOTAL'] > 0) {
                                const fechaFormateada = new Date(fecha).toLocaleDateString('es-ES');
                                const fila = [fechaFormateada];

                                categoriasAMostrar.forEach(categoria => {
                                    fila.push(resumenPorGestor[agente].porFecha[fecha][categoria] || 0);
                                });

                                fila.push(resumenPorGestor[agente].porFecha[fecha]['TOTAL'] || 0);

                                resumenNivel1Sheet.addRow(fila);
                            }
                        });

                        // Fila de totales para este agente
                        const totalesAgenteFila = ['TOTAL AGENTE'];

                        categoriasAMostrar.forEach(categoria => {
                            totalesAgenteFila.push(resumenPorGestor[agente].totales[categoria] || 0);
                        });

                        totalesAgenteFila.push(resumenPorGestor[agente].totales['TOTAL'] || 0);

                        const totalAgenteRow = resumenNivel1Sheet.addRow(totalesAgenteFila);

                        totalAgenteRow.eachCell((cell) => {
                            cell.style = {
                                font: { bold: true },
                                fill: { type: 'pattern', pattern: 'solid', fgColor: { argb: 'D7BDE2' } }
                            };
                        });
                    }
                });

                // Autoajustar columnas
                resumenNivel1Sheet.columns.forEach((column) => {
                    let maxLength = 0;
                    column.eachCell({ includeEmpty: true }, (cell) => {
                        const length = cell.value ? cell.value.toString().length : 10;
                        if (length > maxLength) maxLength = length;
                    });
                    column.width = Math.min(maxLength + 2, 25);
                });

                // Estadísticas de procesamiento
                resumenNivel1Sheet.addRow([]);
                resumenNivel1Sheet.addRow(['ESTADÍSTICAS DE PROCESAMIENTO', '', '', '', '', '', '', '', '']);
                resumenNivel1Sheet.addRow(['Registros procesados:', registrosProcesados, '', '', '', '', '', '', '']);
                resumenNivel1Sheet.addRow(['Registros no clasificados:', registrosNoClasificados, '', '', '', '', '', '', '']);
                resumenNivel1Sheet.addRow(['Porcentaje de éxito:', `${(registrosProcesados / datos.todosRegistros.length * 100).toFixed(2)}%`, '', '', '', '', '', '', '']);

            } else {
                resumenNivel1Sheet.addRow([]);
                resumenNivel1Sheet.addRow(['¡ERROR! No se pudo encontrar la columna "Nivel 1"', '', '', '', '', '', '', '', '']);

                // Mostrar todas las columnas disponibles para debug
                resumenNivel1Sheet.addRow([]);
                resumenNivel1Sheet.addRow(['Columnas disponibles en el archivo:', '', '', '', '', '', '', '', '']);
                columnas.forEach((columna, index) => {
                    resumenNivel1Sheet.addRow([`${index + 1}. ${columna}`, '', '', '', '', '', '', '', '']);
                });
            }
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
        const hojasCount = 6; // Ahora siempre 6 hojas
        a.download = `${nombreBase}_reporte_${hojasCount}hojas_${fechaHoy}.xlsx`;

        document.body.appendChild(a);
        a.click();
        document.body.removeChild(a);
        window.URL.revokeObjectURL(url);

        let mensaje = `Reporte exportado exitosamente con ${hojasCount} hojas (incluye resumen por nivel 1)`;
        if (documentosConflictivos && documentosConflictivos.length > 0) {
            mensaje += ` - ${documentosConflictivos.length} documentos con múltiples agentes por día`;
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
