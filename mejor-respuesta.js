// Procesador de Mejor Respuesta
function procesarMejorRespuesta(data, tipo) {
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

        // Solo filas que tengan un nivel válido en la jerarquía (comparación sin mayúsculas/minúsculas)
        const filasValidas = grupo.filter(row => {
            const nivel = (row[nivelColumna] || '').trim().toUpperCase();
            return jerarquia.some(j => j.toUpperCase() === nivel);
        });

        if (filasValidas.length === 0) return;

        // Calcular índice en la jerarquía para cada fila válida
        const filasConIndice = filasValidas.map(row => {
            const nivel = (row[nivelColumna] || '').trim().toUpperCase();
            const idx = jerarquia.findIndex(j => j.toUpperCase() === nivel);
            return { row, idx };
        });

        const maxIdx = Math.max(...filasConIndice.map(f => f.idx));
        const minIdx = Math.min(...filasConIndice.map(f => f.idx));

        const mejor = jerarquia[maxIdx];
        const peor = jerarquia[minIdx];

        // Fila específica que tiene la mejor gestión → tomar su Comentarios
        const filaMejor = filasConIndice.find(f => f.idx === maxIdx).row;

        // Usar la primera fila del grupo como base de datos generales
        const filaBase = { ...grupo[0] };

        filaBase['Mejor Gestión'] = mejor;
        filaBase['Comentario Mejor Gestión'] = filaMejor['Comentarios'] || '';
        filaBase['Peor Gestión'] = peor;
        filaBase['Total Gestiones'] = grupo.length;

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
