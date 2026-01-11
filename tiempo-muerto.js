// Procesador de Tiempo Muerto
function procesarTiempoMuerto(data, tipo) {
    const nivelColumna = tipo === 'dmiro' ? 'Nivel 2' : 'Nivel 3';
    const fechaColumna = tipo === 'dmiro' ? 'Fecha' : 'Fecha de Gestión';
    const horaColumna = 'HORA';
    const agenteColumna = 'Agente';
    
    const nivelesPermitidos = ['CONTESTADOR AUTOMATICO', 'NO CONTESTA', 'NUMERO EQUIVOCADO'];
    
    // PASO 1: Procesar TODOS los registros
    const todosRegistros = [];
    
    data.forEach(row => {
        const agente = row[agenteColumna];
        if (!agente) return;
        
        let hora;
        
        if (tipo === 'dmiro' && row[horaColumna]) {
            hora = typeof row[horaColumna] === 'string' ? row[horaColumna] :
                  new Date(row[horaColumna]).toTimeString().split(' ')[0];
        } else {
            const fechaCompleta = row[fechaColumna];
            if (!fechaCompleta) return;
            
            if (typeof fechaCompleta === 'string') {
                const partes = fechaCompleta.split(' ');
                hora = partes.length >= 2 ? partes[1] : null;
            } else if (fechaCompleta instanceof Date) {
                hora = fechaCompleta.toTimeString().split(' ')[0];
            } else if (typeof fechaCompleta === 'number') {
                const date = new Date((fechaCompleta - 25569) * 86400 * 1000);
                hora = date.toTimeString().split(' ')[0];
            } else {
                const partes = String(fechaCompleta).split(' ');
                hora = partes.length >= 2 ? partes[1] : null;
            }
        }
        
        if (!hora || !hora.includes(':')) return;
        const [horas, minutos, segundos] = hora.split(':').map(p => parseInt(p) || 0);
        if (horas > 23 || minutos > 59 || (segundos && segundos > 59)) return;
        
        // Excluir registros entre 12:00 y 13:59 (almuerzo)
        if (horas >= 12 && horas < 14) return;
        
        const totalSegundos = horas * 3600 + minutos * 60 + segundos;
        
        todosRegistros.push({
            ...row,
            HORA_EXTRAIDA: hora,
            HORA_NUM: totalSegundos,
            AGENTE_CLEAN: agente,
            ES_NIVEL_PERMITIDO: nivelesPermitidos.includes((row[nivelColumna] || '').trim().toUpperCase())
        });
    });
    
    // PASO 2: Agrupar por agente
    const grupos = {};
    todosRegistros.forEach(row => {
        const agente = row.AGENTE_CLEAN;
        if (!grupos[agente]) grupos[agente] = [];
        grupos[agente].push(row);
    });
    
    const datosFiltrados = [];
    const resumenTiempoMuerto = [];
    
    Object.keys(grupos).forEach(agente => {
        const todosRegistrosAgente = grupos[agente];
        if (todosRegistrosAgente.length <= 1) return;
        
        // Ordenar por hora (mayor a menor)
        todosRegistrosAgente.sort((a, b) => b.HORA_NUM - a.HORA_NUM);
        
        // PASO 3: Calcular diferencias
        for (let i = 0; i < todosRegistrosAgente.length; i++) {
            const registro = todosRegistrosAgente[i];
            
            if (i < todosRegistrosAgente.length - 1) {
                const siguiente = todosRegistrosAgente[i + 1];
                const diferencia = registro.HORA_NUM - siguiente.HORA_NUM;
                
                if (diferencia > 0) {
                    registro.TIEMPO_MUERTO = formatearTiempo(diferencia);
                    registro.TIEMPO_MUERTO_SEGUNDOS = diferencia;
                } else {
                    registro.TIEMPO_MUERTO = 'Excluido - Diferencia inválida';
                    registro.TIEMPO_MUERTO_SEGUNDOS = 0;
                }
            } else {
                registro.TIEMPO_MUERTO = 'N/A - Último registro';
                registro.TIEMPO_MUERTO_SEGUNDOS = 0;
            }
        }
        
        // PASO 4: Filtrar solo niveles permitidos
        const registrosFiltrados = todosRegistrosAgente.filter(reg => reg.ES_NIVEL_PERMITIDO);
        
        // Calcular resumen
        let tiempoMuertoTotalSegundos = 0;
        let intervalosValidos = 0;
        
        registrosFiltrados.forEach(reg => {
            if (reg.TIEMPO_MUERTO_SEGUNDOS > 0) {
                tiempoMuertoTotalSegundos += reg.TIEMPO_MUERTO_SEGUNDOS;
                intervalosValidos++;
            }
        });
        
        // Agregar registros filtrados
        registrosFiltrados.forEach(reg => {
            const filaFinal = { ...reg };
            delete filaFinal.AGENTE_CLEAN;
            delete filaFinal.HORA_NUM;
            delete filaFinal.ES_NIVEL_PERMITIDO;
            datosFiltrados.push(filaFinal);
        });
        
        if (registrosFiltrados.length > 0) {
            resumenTiempoMuerto.push({
                'Agente': agente,
                'Tiempo Muerto Total': formatearTiempo(tiempoMuertoTotalSegundos),
                'Tiempo Muerto Total (Segundos)': tiempoMuertoTotalSegundos,
                'Número de Intervalos Válidos': intervalosValidos,
                'Total Registros Filtrados': registrosFiltrados.length,
                'Total Registros Agente': todosRegistrosAgente.length,
                'Primer Registro': registrosFiltrados[registrosFiltrados.length - 1]?.HORA_EXTRAIDA || '',
                'Último Registro': registrosFiltrados[0]?.HORA_EXTRAIDA || ''
            });
        }
    });
    
    // Ordenar por tiempo muerto total
    resumenTiempoMuerto.sort((a, b) => {
        return b['Tiempo Muerto Total (Segundos)'] - a['Tiempo Muerto Total (Segundos)'];
    });
    
    return {
        tipo: 'tiempo_muerto',
        datosFiltrados: datosFiltrados,
        resumenTiempoMuerto: resumenTiempoMuerto
    };
}

// Exportar a Excel
async function exportarTiempoMuerto(datos, nombreArchivo) {
    const workbook = new ExcelJS.Workbook();
    
    // Hoja 1: Datos Filtrados
    const ws1 = workbook.addWorksheet('Datos Filtrados');
    if (datos.datosFiltrados.length > 0) {
        const headers1 = Object.keys(datos.datosFiltrados[0]);
        ws1.addRow(headers1);
        
        datos.datosFiltrados.forEach(row => {
            ws1.addRow(Object.values(row));
        });
        
        // Autoajuste
        ws1.columns.forEach((column) => {
            let maxLength = 0;
            column.eachCell({ includeEmpty: true }, (cell) => {
                const length = cell.value ? cell.value.toString().length : 10;
                if (length > maxLength) maxLength = length;
            });
            column.width = Math.max(maxLength + 2, 12);
        });
    }
    
    // Hoja 2: Resumen
    const ws2 = workbook.addWorksheet('Resumen Tiempo Muerto');
    
    // Título
    ws2.mergeCells('A1:H1');
    const titleCell = ws2.getCell('A1');
    titleCell.value = 'RESUMEN DE TIEMPO MUERTO POR AGENTE';
    titleCell.style = {
        font: { bold: true, size: 14, color: { argb: 'FFFFFF' } },
        fill: { type: 'pattern', pattern: 'solid', fgColor: { argb: 'E74C3C' } },
        alignment: { horizontal: 'center', vertical: 'middle' }
    };
    
    // Encabezados
    if (datos.resumenTiempoMuerto.length > 0) {
        const headers2 = Object.keys(datos.resumenTiempoMuerto[0]);
        const headerRow = ws2.addRow(headers2);
        
        headerRow.eachCell((cell) => {
            cell.style = {
                font: { bold: true, color: { argb: 'FFFFFF' } },
                fill: { type: 'pattern', pattern: 'solid', fgColor: { argb: '3498DB' } },
                alignment: { horizontal: 'center' }
            };
        });
        
        // Datos
        datos.resumenTiempoMuerto.forEach((item, index) => {
            const fila = headers2.map(h => item[h]);
            const dataRow = ws2.addRow(fila);
            
            // Colores según tiempo muerto
            const tiempoMuerto = item['Tiempo Muerto Total (Segundos)'];
            let fillColor = 'FFFFFF';
            
            if (tiempoMuerto > 3600) fillColor = 'FF6B6B'; // Rojo: más de 1 hora
            else if (tiempoMuerto > 1800) fillColor = 'FFD93D'; // Amarillo: 30-60 min
            else if (tiempoMuerto > 600) fillColor = '6BCF7F'; // Verde: 10-30 min
            
            dataRow.eachCell((cell) => {
                cell.style = {
                    fill: { type: 'pattern', pattern: 'solid', fgColor: { argb: fillColor } },
                    alignment: { horizontal: 'center' }
                };
            });
        });
        
        // Autoajuste
        ws2.columns.forEach((column) => {
            let maxLength = 0;
            column.eachCell({ includeEmpty: true }, (cell) => {
                const length = cell.value ? cell.value.toString().length : 10;
                if (length > maxLength) maxLength = length;
            });
            column.width = Math.max(maxLength + 2, 15);
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
    a.download = nombreArchivo;
    document.body.appendChild(a);
    a.click();
    document.body.removeChild(a);
    window.URL.revokeObjectURL(url);
}

// Exportar al scope global
window.procesarTiempoMuerto = procesarTiempoMuerto;
window.exportarTiempoMuerto = exportarTiempoMuerto;