function seleccionMultip1e(e){
    var celdaActiva=SpreadsheetApp.getActiveRange();
    var columnaActiva=celdaActiva.getColumn()
    var filaActiva=celdaActiva.getRow()
    var valorActivo=e.value
    var valorAnterior=e.oldValue
    var celdaResultado=celdaActiva.offset(0,1)
    var ValorAcumulado=celdaResultado.getValue()

    var busquedaDuplicados=ValorAcumulado.toString().indexOF(valorActivo)

    if(columnaActiva==2 && filaActiva>2 ){
        if(!valorActivo) ValorAcumulado=""
        else if(!valorAnterior) ValorAcumulado=valorActivo
        else if (busquedaDuplicados!=-1) ValorAcumulado=ValorAcumulado
        else if(valorActivo!=valorAnterior){
            ValorAcumulado=ValorAcumulado+","+valorActivo

            }
            celdaResultado.setValue(ValorAcumulado)

        
        
    } 

}

function onEdit(e){
    seleccionMultip1e(e)
}