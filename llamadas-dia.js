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

// Exportar a Excel
async function exportarLlamadasDia(datos, nombreArchivo) {
    try {
        const agentesData = datos.agentesData;
        const fechas = datos.fechas;
        const agentes = Object.keys(agentesData).sort();
        
        // Crear workbook
        const workbook = new ExcelJS.Workbook();
        const worksheet = workbook.addWorksheet('Reporte Llamadas');
        
        // Título
        worksheet.mergeCells('A1:C1');
        const titleCell = worksheet.getCell('A1');
        titleCell.value = 'REPORTE DE LLAMADAS POR DÍA';
        titleCell.style = {
            font: { bold: true, size: 16, color: { argb: 'FFFFFF' } },
            fill: { type: 'pattern', pattern: 'solid', fgColor: { argb: '2C3E50' } },
            alignment: { horizontal: 'center', vertical: 'middle' }
        };
        
        // Estadísticas
        worksheet.addRow([]);
        const statsRow = worksheet.addRow(['Estadísticas Generales', '', '']);
        statsRow.font = { bold: true };
        
        worksheet.addRow(['Total Llamadas', datos.todosRegistros.length, '']);
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
        
        // Datos por agente
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
        
        // Fila de totales
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
        
        // Segunda hoja con detalle
        if (datos.todosRegistros.length > 0) {
            const detalleSheet = workbook.addWorksheet('Detalle');
            
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
            
            // Datos
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
        a.download = `${nombreBase}_reporte_llamadas_${fechaHoy}.xlsx`;
        
        document.body.appendChild(a);
        a.click();
        document.body.removeChild(a);
        window.URL.revokeObjectURL(url);
        
        showStatus('Reporte exportado exitosamente', 'success');
        
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