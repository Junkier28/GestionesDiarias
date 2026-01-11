// Procesador de Número de Gestiones
function procesarNumeroGestiones(data, tipo) {
    const nivelColumna = tipo === 'dmiro' ? 'Nivel 1' : 'Nivel 2';
    const agenteColumna = 'Agente';
    
    const categorias = ['CONTACTO DIRECTO', 'CONTACTO INDIRECTO', 'SIN CONTACTO'];
    
    const resultados = {
        tipo: 'numero_gestiones',
        datosOriginales: data,
        tablaDinamica: {},
        totales: {
            'CONTACTO DIRECTO': 0,
            'CONTACTO INDIRECTO': 0,
            'SIN CONTACTO': 0,
            'Total general': 0
        }
    };
    
    // Procesar cada fila
    data.forEach(row => {
        const agente = row[agenteColumna];
        const nivel = row[nivelColumna];
        
        if (!agente) return;
        
        // Inicializar agente
        if (!resultados.tablaDinamica[agente]) {
            resultados.tablaDinamica[agente] = {
                'CONTACTO DIRECTO': 0,
                'CONTACTO INDIRECTO': 0,
                'SIN CONTACTO': 0,
                'Total general': 0
            };
        }
        
        // Determinar categoría
        let categoria = 'SIN CONTACTO';
        
        if (nivel) {
            const nivelUpper = nivel.toUpperCase();
            
            if (nivelUpper.includes('CONTACTO DIRECTO')) {
                categoria = 'CONTACTO DIRECTO';
            } else if (nivelUpper.includes('CONTACTO INDIRECTO')) {
                categoria = 'CONTACTO INDIRECTO';
            }
        }
        
        // Incrementar contadores
        resultados.tablaDinamica[agente][categoria]++;
        resultados.tablaDinamica[agente]['Total general']++;
        
        resultados.totales[categoria]++;
        resultados.totales['Total general']++;
    });
    
    return resultados;
}

// Exportar a Excel
async function exportarNumeroGestiones(datos, nombreArchivo) {
    const workbook = new ExcelJS.Workbook();
    
    // Hoja 1: Datos Originales
    if (datos.datosOriginales.length > 0) {
        const worksheet1 = workbook.addWorksheet('Datos Originales');
        const headers1 = Object.keys(datos.datosOriginales[0]);
        worksheet1.addRow(headers1);
        
        datos.datosOriginales.forEach(row => {
            worksheet1.addRow(Object.values(row));
        });
        
        // Autoajuste
        worksheet1.columns.forEach((column) => {
            let maxLength = 0;
            column.eachCell({ includeEmpty: true }, (cell) => {
                const length = cell.value ? cell.value.toString().length : 10;
                if (length > maxLength) maxLength = length;
            });
            column.width = Math.max(maxLength + 2, 12);
        });
    }
    
    // Hoja 2: Resumen por Agente
    const worksheet2 = workbook.addWorksheet('Resumen por Agente');
    
    // Título
    worksheet2.mergeCells('A1:E1');
    const titleCell = worksheet2.getCell('A1');
    titleCell.value = 'RESUMEN DE GESTIONES POR AGENTE';
    titleCell.style = {
        font: { bold: true, size: 14, color: { argb: 'FFFFFF' } },
        fill: { type: 'pattern', pattern: 'solid', fgColor: { argb: '2F75B5' } },
        alignment: { horizontal: 'center', vertical: 'middle' }
    };
    
    // Encabezados
    const headers = ['Agente', 'CONTACTO DIRECTO', 'CONTACTO INDIRECTO', 'SIN CONTACTO', 'Total general'];
    worksheet2.addRow([]);
    const headerRow = worksheet2.addRow(headers);
    
    headerRow.eachCell((cell) => {
        cell.style = {
            font: { bold: true, color: { argb: 'FFFFFF' } },
            fill: { type: 'pattern', pattern: 'solid', fgColor: { argb: '5B9BD5' } },
            alignment: { horizontal: 'center', vertical: 'middle' }
        };
    });
    
    // Datos por agente
    const agentes = Object.keys(datos.tablaDinamica);
    agentes.forEach((agente, index) => {
        const fila = [
            agente,
            datos.tablaDinamica[agente]['CONTACTO DIRECTO'],
            datos.tablaDinamica[agente]['CONTACTO INDIRECTO'],
            datos.tablaDinamica[agente]['SIN CONTACTO'],
            datos.tablaDinamica[agente]['Total general']
        ];
        
        const dataRow = worksheet2.addRow(fila);
        
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
    worksheet2.addRow([]);
    const totalRow = worksheet2.addRow([
        'TOTAL GENERAL',
        datos.totales['CONTACTO DIRECTO'],
        datos.totales['CONTACTO INDIRECTO'],
        datos.totales['SIN CONTACTO'],
        datos.totales['Total general']
    ]);
    
    totalRow.eachCell((cell) => {
        cell.style = {
            font: { bold: true, color: { argb: 'FFFFFF' } },
            fill: { type: 'pattern', pattern: 'solid', fgColor: { argb: 'ED7D31' } },
            alignment: { horizontal: 'center' }
        };
    });
    
    // Autoajuste
    worksheet2.columns.forEach((column) => {
        let maxLength = 0;
        column.eachCell({ includeEmpty: true }, (cell) => {
            const length = cell.value ? cell.value.toString().length : 10;
            if (length > maxLength) maxLength = length;
        });
        column.width = Math.max(maxLength + 2, 12);
    });
    
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
window.procesarNumeroGestiones = procesarNumeroGestiones;
window.exportarNumeroGestiones = exportarNumeroGestiones;