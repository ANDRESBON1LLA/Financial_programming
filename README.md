# Financial_programming
En el siguiente documento, vamos a crear una base de datos de acitvos financieros y simular su comportamiento para construir una tabla con dicho contenido. Finalmente convertir esta tabla en una tabla dinámica y generar gráficos interactivos.

## Generación de datos sobre activos financieros y commodities.

El siguiente código genera los datos sobre los siguientes activos financieros y activos tangibles (Identificados en la tabla de datos por sus nemotecnicos): <br>
&nbsp;&nbsp;&nbsp;&nbsp; -Activos representativos de participación (Acciones): Apple, Microsoft, Google, Amazon, Tesla. <br>
&nbsp;&nbsp;&nbsp;&nbsp; -Criptoactivos: Bitcoin, Ethereum, XRP. <br>
&nbsp;&nbsp;&nbsp;&nbsp; -Commodities: Oro, Plata, Petróleo. <br>
&nbsp;&nbsp;&nbsp;&nbsp; -Activos de deuda: Bonos de USA
  
```vba

Sub GenerarDatosActivosFinancieros2()
    Dim ws As Worksheet 'La variable ws definida como hoja de trabajo
    Dim tbl As ListObject 'La variable tbl definida como una lista de objetos
    Dim rng As Range 'La variable rng definida como un rango
    Dim i As Integer 'La variable i definida como un valor entero
    Dim activos As Variant 'La variable activos definido como variable
    Dim filaInicio As Integer 'La variable filaInicio definida como un valor entero
    Dim numDatos As Integer 'La variable numDatos definida como valor entero
    Dim preciosBase As Variant 'La variable preciosBase se utiliza para definir los precios iniciales de los activos
    Dim volatilidad As Variant
    Dim precioActual As Double 'La variable precioActual suscribe el nuevo precio según el comportamiento de la volatilidad
    Dim variacion As Double
    Dim cantidad As Double
    Dim valorTotal As Double
    Dim claseActivo As String
    
    
    ' Definir nombres de activos financieros y sus valores iniciales
    activos = Array("AAPL", "MSFT", "GOOGL", "AMZN", "TSLA", "BTC", "ETH", "XRP", "Oro", "Plata", "Petróleo", "Bonos USA")
    
    ' Precios base para cada activo
    preciosBase = Array(175, 320, 2800, 3500, 750, 45000, 3200, 1.1, 1900, 25, 85, 100) ' Valores aproximados

    ' Volatilidad esperada (% diario máximo)
    volatilidad = Array(1, 1.2, 1.5, 2, 3, 5, 6, 8, 0.8, 0.6, 1.5, 0.3)

    ' Número de registros a generar por activo
    numDatos = 50 ' Puedes cambiar la cantidad de días a simular

    ' Crear hoja de datos si no existe
    On Error Resume Next
    Set ws = Worksheets("Portafolio")
    On Error GoTo 0
    If ws Is Nothing Then
        Set ws = Worksheets.Add
        ws.Name = "Portafolio"
    End If
    
    ' Borrar datos anteriores
    ws.Cells.Clear
    
    ' Definir encabezados de la tabla
    ws.Range("A1:I1").Value = Array("Activo", "Clase de activo", "Cantidad", "Precio", "Variación %", "Volumen", "Capitalización", "Fecha", "Valor Total")
    
    ' Inicializar valores para la simulación
    filaInicio = 2 'Desde la fila 2 se ubicarán los valores
    Dim j As Integer
    Dim filaActual As Integer ' Variable para manejar la fila correctamente
    
    filaActual = filaInicio ' La fila donde empezaremos a escribir
    
    For j = LBound(activos) To UBound(activos) ' Recorre cada activo
        
        ' Inicializar el precio con su valor base
        precioActual = preciosBase(j)
        
        ' Determinar la cantidad y la clase de activo según el tipo de activo
        Select Case activos(j) 'Select case
            Case "AAPL", "MSFT", "GOOGL", "AMZN", "TSLA" ' Acciones
                cantidad = Int(10 + Rnd * 490) ' 10 a 500 acciones. IMPORTANTE INT YA QUE NO EXISTEN ACCIONES CON CIFRAS DECIMALES
                claseActivo = "Acciones"
            Case "BTC", "ETH", "XRP" ' Criptomonedas
                cantidad = Round(0.01 + Rnd * 4.99, 2) ' 0.01 a 5 monedas
                claseActivo = "Criptomonedas"
            Case "Oro", "Plata", "Petróleo" ' Commodities
                cantidad = Int(1 + Rnd * 99) ' 1 a 100 contratos
                claseActivo = "Commodities"
            Case "Bonos USA" ' Bonos
                cantidad = Int(1 + Rnd * 49) ' 1 a 50 bonos
                claseActivo = "Bonos"
            Case Else
                cantidad = 1 ' Valor por defecto
                claseActivo = "Otros"
        End Select
        
        
        ' Determinar la cantidad según el tipo de activo
        
        ' Generar datos estructurados para cada activo
        For i = 0 To numDatos - 1
        'Se utiliza desde 0 hasta 49 dando un total de 50 datos
            ' Variación diaria basada en volatilidad propia
            variacion = Round((-volatilidad(j) + Rnd * (volatilidad(j) * 2)), 2) ' Variación entre -X% y +X%
            
            'El Rnd selecciona un valor aleatorio entre 0-1
            'Round redondea el valor de la volatilidad utilizando 2 decimales (Expuesto en el segundo argumento).
            'La volatilidad se multiplica por 2 para garantizar un intervalo entre valores negativos y positivos.
            'EJ: si Rnd=0 -> -3+0=-3.
            'Si Rnd=0.5-> -3+3=0
            'Si Rnd=1 -> -3+6=3
            
            ' Actualizar el precio con base en la variación
            precioActual = Round(precioActual * (1 + variacion / 100), 2) 'Acá se aplica una actualización del precio de manera porcentual (1.01, 0.98, 0.90)
            
                       ' Calcular el valor total de la posición
            valorTotal = Round(cantidad * precioActual, 2)
            
            ' Insertar datos en la hoja en la fila correspondiente
            ws.Cells(filaActual, 1).Value = activos(j) ' Activo
            'El codigo se define de la siguiente manera: FilaInicio+i define la fila en qu ese ubica.
            'El segundo termino define la columna en que se va a ubicar el activo (1, 2, 3, 4)
            ws.Cells(filaActual, 2).Value = claseActivo ' Clase de activo
            ws.Cells(filaActual, 3).Value = cantidad ' Cantidad en el portafolio
            ws.Cells(filaActual, 4).Value = precioActual ' Precio actualizado
            ws.Cells(filaActual, 5).Value = variacion ' Variación porcentual
            ws.Cells(filaActual, 6).Value = Round(500000 + (Rnd * 1000000), 0) ' Volumen con sentido lógico
            ws.Cells(filaActual, 7).Value = Round(precioActual * 1000000, 0) ' Capitalización aproximada
            ws.Cells(filaActual, 8).Value = Date - (numDatos - i) ' Fechas en orden cronológico
            ws.Cells(filaActual, 9).Value = valorTotal ' Valor total de la posición
            
            ' Mover a la siguiente fila
            filaActual = filaActual + 1
        Next i
    Next j

    ' Convertir el rango en una tabla si aún no existe
    Set rng = ws.Range("A1").CurrentRegion 'La función current region selecciona todas las celdas adyacentes con datos.
    'Ej: Si tienes datos desde A1 hasta F50, CurrentRegion seleccionará automáticamente todo ese bloque.
    On Error Resume Next
    Set tbl = ws.ListObjects("Tabla1")
    On Error GoTo 0
    If tbl Is Nothing Then
        Set tbl = ws.ListObjects.Add(xlSrcRange, rng, , xlYes)
        tbl.Name = "Tabla1"
    End If

    MsgBox "Portafolio de activos generado", vbInformation
End Sub
```
Revisemos que se creó correctamente la sheet con el nombre de "Portafolio" y que de manera simultánea, la tabla recibió el nombre de "Tabla1". Se debe ejecutar además un mensaje que confirme la generación de datos con coherencia.

Es importante que la tabla se genere correctamente, de otra manera no habrá el insumo necesario para la generación de la tabla dinámica.

## Ahora debemos crear la tabla dinámica:

Antes eliminará una hoja llamada tabla dinámica si esta ya existe, para dar paso a la nueva hoja titulada de la misma manera

```vba
Sub CrearTablaDinamica()
    Dim ws As Worksheet 'Definir ws como una hoja de trabajo
    Dim PCache As PivotCache 'Mas adelante se define que es un Pivot Cache
    Dim TDinamica As PivotTable
    Dim wsName As String
    wsName = "TablaDinamica" 'O el nombre que le deseemos dar a la hoja
    ' Verificar si la hoja "TablaDinamica" existe
    On Error Resume Next 'Modo de error silencioso: No generará un error al intentar asignarlo a ws
    Set ws = Worksheets(wsName) 'Si arroja error: ws=Nothing
    On Error GoTo 0 'Reactiva el manejo de errores para las siguientes lineas
    ' Si la hoja no existe, la creamos
    If ws Is Nothing Then
        Set ws = Worksheets.Add
        ws.Name = wsName
    Else
        ' Si existe, limpiamos la hoja en lugar de eliminarla
        ws.Cells.Clear
    End If
    ' Crear Pivot Cache (almacenará los datos de la tabla para la tabla dinámica)
    On Error Resume Next
    Set PCache = ActiveWorkbook.PivotCaches.Create(xlDatabase, "Tabla1")
    On Error GoTo 0
    ' Verificar si el Pivot Cache se creó correctamente, o si la tabla 1 no existe
    If PCache Is Nothing Then
        MsgBox "Error: La tabla 'Tabla1' no existe o los datos no están en un formato adecuado.", vbCritical
        Exit Sub
    End If
    ' Crear la tabla dinámica en la celda C3 de "TablaDinamica"
    Set TDinamica = PCache.CreatePivotTable(ws.Range("C3"), "Tabla dinámica")
    'Primer argumento: Indica la ubicación de la nueva tabla dinámica.
    'El .Range("C3") indica la casilla particular
    ' Confirmación de éxito
    MsgBox "Tabla dinámica creada correctamente en la hoja '" & wsName & "'.", vbInformation
End Sub
```
Revisar que la tabla dinámica exista en la celda C3 tal como se especificó en los argumentos.


