// Procesador de Mejor Respuesta
function procesarMejorRespuesta(data, tipo) {
    const jerarquia = config.jerarquias[tipo];
    const nivelColumna = tipo === 'dmiro' ? 'Nivel 2' : 'Nivel 3';
    const idColumna = tipo === 'dmiro' ? 'Documento' : 'Número Credito';
    
    console.log('=== MEJOR RESPUESTA DEBUG ===');
    console.log('TIPO:', tipo);
    console.log('NIVEL COLUMNA:', nivelColumna);
    console.log('ID COLUMNA:', idColumna);
    console.log('PRIMERA FILA COMPLETA:', JSON.stringify(data[0]));
    console.log('COLUMNAS DISPONIBLES:', Object.keys(data[0]));
    console.log('JERARQUIA:', jerarquia);
    
    const grupos = {};
    
    // Agrupar por ID
    data.forEach(row => {
        const id = row[idColumna];
        if (!id) return;
        
        if (!grupos[id]) grupos[id] = [];
        grupos[id].push(row);
    });
    
    console.log('TOTAL GRUPOS (IDs únicos):', Object.keys(grupos).length);
    
    // Mostrar el primer grupo como ejemplo
    const primerId = Object.keys(grupos)[0];
    if (primerId) {
        console.log('EJEMPLO - ID:', primerId);
        console.log('EJEMPLO - FILAS:', grupos[primerId].length);
        grupos[primerId].forEach((row, i) => {
            console.log(`EJEMPLO - FILA ${i}: Nivel=${row[nivelColumna]}, Comentarios=${row['Comentarios']}`);
        });
    }
    
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
            
            // Buscar la fila exacta que tiene la mejor gestión
            const filaMejor = grupo.find(row => row[nivelColumna] === mejor);
            
            console.log(`ID ${id} -> Mejor: ${mejor}, FilaMejor encontrada: ${!!filaMejor}, Comentarios: ${filaMejor ? filaMejor['Comentarios'] : 'NO ENCONTRADA'}`);
            
            const filaEjemplo = { ...grupo[0] };
            filaEjemplo['Mejor Gestión'] = mejor;
            filaEjemplo['Comentario Mejor Gestión'] = filaMejor ? (filaMejor['Comentarios'] || '') : '';
            filaEjemplo['Peor Gestión'] = peor;
            filaEjemplo['Total Gestiones'] = grupo.length;
            
            resultado.push(filaEjemplo);
        }
    });
    
    console.log('TOTAL RESULTADO:', resultado.length);
    console.log('=== FIN DEBUG ===');
    
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
