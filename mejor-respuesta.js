// Procesador de Mejor Respuesta
function procesarMejorRespuesta(data, tipo) {
    const jerarquia = config.jerarquias[tipo];
    const nivelColumna = tipo === 'dmiro' ? 'Nivel 2' : 'Nivel 3';
    const idColumna = tipo === 'dmiro' ? 'Documento' : 'Número Credito';

    const grupos = {};

    // Agrupar por ID — forzar string para evitar pérdida de precisión en números largos
    data.forEach(row => {
        const id = row[idColumna] !== undefined && row[idColumna] !== null
            ? String(row[idColumna]).trim()
            : null;
        if (!id) return;
        if (!grupos[id]) grupos[id] = [];
        grupos[id].push(row);
    });

    const resultado = [];

    Object.keys(grupos).forEach(id => {
        const grupo = grupos[id];

        // Mapear cada fila a su índice en la jerarquía (-1 si no está)
        const filasConIndice = grupo.map(row => {
            const nivel = (row[nivelColumna] || '').trim().toUpperCase();
            const idx = jerarquia.findIndex(j => j.trim().toUpperCase() === nivel);
            return { row, idx };
        }).filter(f => f.idx !== -1);

        if (filasConIndice.length === 0) return;

        const maxIdx = Math.max(...filasConIndice.map(f => f.idx));
        const minIdx = Math.min(...filasConIndice.map(f => f.idx));

        // Fila exacta de la MEJOR gestión → tomar SU comentario
        const filaMejor = filasConIndice.find(f => f.idx === maxIdx).row;

        const filaBase = { ...filaMejor };

        // Forzar el ID como string en el resultado también
        filaBase[idColumna] = id;

        filaBase['Mejor Gestión']            = jerarquia[maxIdx];
        filaBase['Comentario Mejor Gestión'] = filaMejor['Comentarios'] || '';
        filaBase['Peor Gestión']             = jerarquia[minIdx];
        filaBase['Total Gestiones']          = grupo.length;

        resultado.push(filaBase);
    });

    return resultado;
}

// Exportar a Excel — forzar columna ID como texto
async function exportarMejorRespuesta(datos, nombreArchivo) {
    const wb = XLSX.utils.book_new();
    const ws = XLSX.utils.json_to_sheet(datos);

    // Forzar formato texto en la columna Número Credito / Documento
    // para que Excel no trunque números largos
    const range = XLSX.utils.decode_range(ws['!ref']);
    for (let C = range.s.c; C <= range.e.c; C++) {
        const headerCell = ws[XLSX.utils.encode_cell({ r: 0, c: C })];
        if (!headerCell) continue;
        if (headerCell.v === 'Número Credito' || headerCell.v === 'Documento') {
            for (let R = 1; R <= range.e.r; R++) {
                const cellAddr = XLSX.utils.encode_cell({ r: R, c: C });
                if (ws[cellAddr]) {
                    ws[cellAddr].t = 's'; // forzar tipo string
                    ws[cellAddr].v = String(ws[cellAddr].v);
                }
            }
        }
    }

    XLSX.utils.book_append_sheet(wb, ws, "Mejor Respuesta");
    XLSX.writeFile(wb, nombreArchivo);
}

// Exportar al scope global
window.procesarMejorRespuesta = procesarMejorRespuesta;
window.exportarMejorRespuesta = exportarMejorRespuesta;
