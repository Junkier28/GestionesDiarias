// Procesador de Consolidar Comentarios
function procesarConsolidarComentarios(data, tipo) {
    // Solo funciona para Pacífico
    if (tipo === 'dmiro') {
        showStatus('Esta función solo está disponible para archivos de Pacífico', 'error');
        return [];
    }
    
    const idColumna = 'Número Credito';
    const comentariosColumna = 'Comentarios';
    
    // Agrupar por Número Crédito
    const grupos = {};
    
    data.forEach(row => {
        const numeroCredito = row[idColumna];
        const comentario = row[comentariosColumna];
        
        if (!numeroCredito) return;
        
        // Inicializar grupo si no existe
        if (!grupos[numeroCredito]) {
            grupos[numeroCredito] = {
                primeraFila: { ...row },
                comentarios: []
            };
        }
        
        // Agregar comentario si existe
        if (comentario && comentario.toString().trim() !== '') {
            grupos[numeroCredito].comentarios.push(comentario.toString().trim());
        }
    });
    
    // Crear resultado consolidado
    const resultado = [];
    
    Object.keys(grupos).forEach(numeroCredito => {
        const grupo = grupos[numeroCredito];
        const filaConsolidada = { ...grupo.primeraFila };
        
        // Unir comentarios con punto y coma
        if (grupo.comentarios.length > 0) {
            filaConsolidada[comentariosColumna] = grupo.comentarios.join('; ');
        } else {
            filaConsolidada[comentariosColumna] = '';
        }
        
        resultado.push(filaConsolidada);
    });
    
    return resultado;
}

// Exportar a Excel
async function exportarConsolidarComentarios(datos, nombreArchivo) {
    const wb = XLSX.utils.book_new();
    const ws = XLSX.utils.json_to_sheet(datos);
    XLSX.utils.book_append_sheet(wb, ws, "Comentarios Consolidados");
    XLSX.writeFile(wb, nombreArchivo);
}

// Exportar al scope global
window.procesarConsolidarComentarios = procesarConsolidarComentarios;
window.exportarConsolidarComentarios = exportarConsolidarComentarios;