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
        const nivelesVal = grupo
            .map(row => row[nivelColumna])
            .filter(val => val && jerarquia.includes(val));
        
        if (nivelesVal.length > 0) {
            const indices = nivelesVal.map(val => jerarquia.indexOf(val));
            const minIndex = Math.min(...indices);
            const maxIndex = Math.max(...indices);
            
            const peor = jerarquia[minIndex];
            const mejor = jerarquia[maxIndex];
            
            const filaEjemplo = { ...grupo[0] };
            filaEjemplo['Mejor Gestión'] = mejor;
            filaEjemplo['Peor Gestión'] = peor;
            filaEjemplo['Total Gestiones'] = grupo.length;
            
            resultado.push(filaEjemplo);
        }
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