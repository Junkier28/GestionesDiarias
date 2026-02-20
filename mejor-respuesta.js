// Procesador de Mejor Respuesta
function procesarMejorRespuesta(data, tipo) {
        console.log('TIPO:', tipo);
    console.log('PRIMERA FILA:', JSON.stringify(data[0]));
    console.log('JERARQUIA:', config.jerarquias[tipo]);
    const jerarquia = config.jerarquias[tipo];
    const nivelColumna = tipo === 'dmiro' ? 'Nivel 2' : 'Nivel 3';
    const idColumna = tipo === 'dmiro' ? 'Documento' : 'Número Credito';

    const grupos = {};

    // Agrupar por ID
    data.forEach(row => {
        const id = row[idColumna];
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
        }).filter(f => f.idx !== -1); // solo filas con nivel válido

        if (filasConIndice.length === 0) return;

        const maxIdx = Math.max(...filasConIndice.map(f => f.idx));
        const minIdx = Math.min(...filasConIndice.map(f => f.idx));

        // Fila exacta de la MEJOR gestión → para tomar SU comentario
        const filaMejor = filasConIndice.find(f => f.idx === maxIdx).row;
        const filaPeor  = filasConIndice.find(f => f.idx === minIdx).row;

        // Base: usamos la fila de la mejor gestión como referencia principal
        const filaBase = { ...filaMejor };

        filaBase['Mejor Gestión']            = jerarquia[maxIdx];
        filaBase['Comentario Mejor Gestión'] = filaMejor['Comentarios'] || '';
        filaBase['Peor Gestión']             = jerarquia[minIdx];
        filaBase['Total Gestiones']          = grupo.length;

        resultado.push(filaBase);
    });

    return resultado;
}

// Exportar a Excel
async function exportarMejorRespuesta(datos, nombreArchivo) {
    const wb = XLSX.utils.book_new();
    const ws = XLSX.utils.json_to_sheet(datos);
    XLSX.utils.book_append_sheet(wb, ws, "Mejor Respuesta");
    XLSX.writeFile(wb, nombreArchivo);
}

// Exportar al scope global
window.procesarMejorRespuesta = procesarMejorRespuesta;
window.exportarMejorRespuesta = exportarMejorRespuesta;

