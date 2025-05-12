Attribute VB_Name = "KPIs"
Sub Macro()
    Todas_Preguntas
    Ola
    Base_todas_preguntas
    DifSig
End Sub


Sub Todas_Preguntas()
'Se definen las variables a utilizar
    Dim wsOrigen As Worksheet
    Dim wsDestino As Worksheet
    Dim ultimaFila As Long
    Dim i As Long, j As Long
    Dim colDestino As Long
    Dim filaDestino As Long
    Dim nombresABuscar As Variant
    Dim preguntasABuscar As Variant
    Dim nombreActual As Variant
    Dim preguntaActual As String
    Dim filaEncontrada As Long
    Dim filaIniciales As Variant
    Dim indexPregunta As Long
    Dim nombreVariantes As Object
    Dim varianteActual As Variant
    

    ' Define las hojas de origen y destino
    Set wsOrigen = ThisWorkbook.Sheets("Datos") 'Cambia "Datos" por el nombre de tu hoja de origen
    Set wsDestino = ThisWorkbook.Sheets("KPIs") 'Cambia "KPIs" por el nombre de tu hoja de destino

    ' Encuentra la última fila con datos en la hoja de origen
    ultimaFila = wsOrigen.Cells(wsOrigen.Rows.Count, 1).End(xlUp).Row

    ' Define las preguntas y las marcas a buscar
    preguntasABuscar = Array( _
        "Pregunta 1", _
        "Pregunta 2", _
        "Pregunta 3", _
        "Pregunta 4", _
        "Pregunta 5")

    nombresABuscar = Array("Movistar", "Vodafone", "Orange", "Jazztel", "Yoigo", _
                           "DIGI mobil", "Más movil", "Pepephone", "Lowi.es", _
                           "O2", "Simyo", "Fi Network")

    ' Define una lista de variaciones para ciertos nombres
    Set nombreVariantes = CreateObject("Scripting.Dictionary")
    nombreVariantes.Add "Más movil", Array("Más movil", "MásMóvil")
    nombreVariantes.Add "Movistar", Array("Movistar", "Movistar/Telefónica")
    nombreVariantes.Add "Vodafone", Array("Vodafone", "Vodafone/Ono")

    ' Define las filas de inicio para cada pregunta
    filaIniciales = Array(4, 21, 38, 55, 72)

    ' Encuentra la primera columna vacía en la hoja de destino
    colDestino = wsDestino.Cells(4, wsDestino.Columns.Count).End(xlToLeft).Column + 1

    ' Procesa cada pregunta
    For indexPregunta = LBound(preguntasABuscar) To UBound(preguntasABuscar)
        preguntaActual = preguntasABuscar(indexPregunta)
        filaDestino = filaIniciales(indexPregunta)

        ' Busca la pregunta actual en la columna A
        For i = 1 To ultimaFila
            If Trim(wsOrigen.Cells(i, 1).Value) = Trim(preguntaActual) Then
            
                ' Una vez encontrada la pregunta, busca los nombres de las marcas
                For Each nombreActual In nombresABuscar
                    filaEncontrada = 0

                    ' Manejo de variaciones para ciertos nombres
                    If nombreVariantes.exists(nombreActual) Then
                        
                        ' Si el nombre tiene variaciones, busca cada una en las primeras 30 filas después de la pregunta
                        ' Se ponen 30 filas para evitar que pueda pasar a la siguiente pregunta y coger datos erróneos
                        For Each varianteActual In nombreVariantes(nombreActual)
                            For j = i + 1 To WorksheetFunction.Min(i + 30, ultimaFila)
                                If StrComp(Trim(wsOrigen.Cells(j, 1).Value), Trim(varianteActual), vbTextCompare) = 0 Then
                                    filaEncontrada = j
                                    Exit For
                                End If
                            Next j
                            If filaEncontrada > 0 Then Exit For
                        Next varianteActual
                    Else
                    
                        ' Si no tiene variaciones, busca el nombre en las primeras 30 filas después de la pregunta
                        ' Se ponen 30 filas para evitar que pueda pasar a la siguiente pregunta y coger datos erróneos
                        For j = i + 1 To ultimaFila
                            If StrComp(Trim(wsOrigen.Cells(j, 1).Value), Trim(nombreActual), vbTextCompare) = 0 Then
                                filaEncontrada = j
                                Exit For
                            End If
                        Next j
                    End If

                    ' Si se encuentra el nombre, copia el valor correspondiente
                    If filaEncontrada > 0 Then
                        wsDestino.Cells(filaDestino, colDestino).Value = (wsOrigen.Cells(filaEncontrada, 2).Value)
                        filaDestino = filaDestino + 1
                    End If
                Next nombreActual
                Exit For
            End If
        Next i
    Next indexPregunta
End Sub


Sub Ola()
    Dim wsOrigen As Worksheet
    Dim wsDestino As Worksheet
    Dim ultimaFila As Long
    Dim i As Long
    Dim colDestino As Integer
    Dim filasDestino() As Variant
    Dim filaDestino As Integer

    ' Define las hojas de origen y destino
    Set wsOrigen = ThisWorkbook.Sheets("Datos")   ' Cambia "Datos" por el nombre de tu hoja de origen
    Set wsDestino = ThisWorkbook.Sheets("KPIs")  ' Cambia "KPIs" por el nombre de tu hoja de destino

    ' Define las filas donde se colocará "OLA"
    filasDestino = Array(2, 19, 36, 53, 70)

    ' Encuentra la última fila con datos en la hoja de origen
    ultimaFila = wsOrigen.Cells(wsOrigen.Rows.Count, 1).End(xlUp).Row
    
    ' Encuentra la primera columna vacía en la hoja de destino
    colDestino = wsDestino.Cells(3, wsDestino.Columns.Count).End(xlToLeft).Column + 1

    ' Busca la fila donde está "OLA" en la hoja de origen
    For i = 1 To ultimaFila
        If Trim(wsOrigen.Cells(i, 1).Value) = "OLA" Then
            ' Si se encuentra "OLA", coloca los valores en las filas específicas del destino
            Dim k As Long
            For k = LBound(filasDestino) To UBound(filasDestino)
                filaDestino = filasDestino(k)
                wsDestino.Cells(filaDestino, colDestino).Value = wsOrigen.Cells(i + 2, 1).Value
            Next k
            Exit For ' Salimos del bucle después de encontrar "OLA"
        End If
    Next i
End Sub


Sub Base_todas_preguntas()
    Dim wsOrigen As Worksheet
    Dim wsDestino As Worksheet
    Dim ultimaFila As Long
    Dim i As Long, j As Long
    Dim colDestino As Integer
    Dim filaDestino As Variant
    Dim preguntasABuscar As Variant
    Dim preguntaActual As String
    Dim encontrado As Boolean
    Dim k As Long

    ' Define las hojas de origen y destino
    Set wsOrigen = ThisWorkbook.Sheets("Datos")   ' Cambia "Datos" si tu hoja tiene otro nombre
    Set wsDestino = ThisWorkbook.Sheets("KPIs")  ' Cambia "KPIs" si tu hoja tiene otro nombre

    ' Encuentra la última fila con datos en la hoja de origen
    ultimaFila = wsOrigen.Cells(wsOrigen.Rows.Count, 1).End(xlUp).Row

    ' Define las preguntas a buscar
    preguntasABuscar = Array( _
        "Pregunta 1", _
        "Pregunta 2", _
        "Pregunta 3", _
        "Pregunta 4", _
        "Pregunta 5")

    ' Define las filas de destino para cada pregunta
    filaDestino = Array(3, 20, 37, 54, 71)

    ' Encuentra la primera columna vacía en la fila de destino
    colDestino = wsDestino.Cells(3, wsDestino.Columns.Count).End(xlToLeft).Column + 1

    ' Procesa cada pregunta
    For i = LBound(preguntasABuscar) To UBound(preguntasABuscar)
        preguntaActual = preguntasABuscar(i)
        encontrado = False

        ' Busca la pregunta en la columna A de la hoja de origen
        For j = 1 To ultimaFila
            If Trim(wsOrigen.Cells(j, 1).Value) = Trim(preguntaActual) Then
                ' Una vez encontrada la pregunta, busca "Registros:" dentro de las siguientes filas
                For k = j + 1 To ultimaFila
                    If InStr(1, Trim(wsOrigen.Cells(k, 1).Value), "Registros:", vbTextCompare) > 0 Then
                        ' Si encuentra "Registros:", toma el valor de la columna 2 y lo guarda en la fila destino correspondiente
                        wsDestino.Cells(filaDestino(i), colDestino).Value = wsOrigen.Cells(k, 2).Value
                        encontrado = True
                        Exit For
                    End If
                Next k
                Exit For
            End If
        Next j
    Next i
End Sub


Sub DifSig()
'Se crean las variables

    Dim ws As Worksheet
    Dim ultimaColumna As Long
    Dim penultimaColumna As Long
    Dim ultimaFila As Long
    Dim fila As Long
    Dim valorDif As Double
    Dim n1 As Double, p1 As Double, n2 As Double, p2 As Double, pooled As Double, se As Double
    Dim baseActualFila As Long

    ' Define la hoja donde se realizará el cálculo
    Set ws = ThisWorkbook.Sheets("KPIs") ' Cambia "KPIs" por el nombre de tu hoja

    ' Encuentra la última columna con datos
    ultimaColumna = ws.Cells(3, ws.Columns.Count).End(xlToLeft).Column
    penultimaColumna = ultimaColumna - 1 ' Columna inmediatamente anterior

    ' Encuentra la última fila con datos (en la última columna)
    ultimaFila = ws.Cells(ws.Rows.Count, ultimaColumna).End(xlUp).Row

    ' Verifica que haya al menos dos columnas con datos
    If penultimaColumna < 1 Then
        MsgBox "No hay suficientes columnas para realizar el cálculo.", vbExclamation
        Exit Sub
    End If

    ' Inicializa la base en la fila 3
    baseActualFila = 3

    ' Recorre las filas para calcular la diferencia significativa
    For fila = 4 To ultimaFila
        ' Verifica si se encuentra la palabra "Ola" en la penúltima columna
        If InStr(1, ws.Cells(fila, penultimaColumna).Value, "Ola", vbTextCompare) > 0 Then
            ' Actualiza la base actual a la celda inmediatamente debajo de "Ola"
            baseActualFila = fila + 1
        End If

        ' Verifica que las celdas contengan valores numéricos antes de procesarlas
        If IsNumeric(ws.Cells(fila, penultimaColumna).Value) And IsNumeric(ws.Cells(fila, ultimaColumna).Value) Then
            ' Lee los valores necesarios para la fórmula
            n1 = ws.Cells(baseActualFila, penultimaColumna).Value ' Tamaño muestral de la base
            p1 = ws.Cells(fila, penultimaColumna).Value / 100 ' Proporción de la base
            n2 = ws.Cells(baseActualFila, ultimaColumna).Value ' Tamaño muestral de la última columna
            p2 = ws.Cells(fila, ultimaColumna).Value / 100 ' Proporción de la última columna

            ' Evita divisiones por cero y valores inválidos para pooled
            If n1 > 0 And n2 > 0 Then
            
                'Se calcula la estimación combinada
                pooled = ((n2 * p2) + (n1 * p1)) / (n2 + n1)
                If pooled > 0 And pooled < 1 Then ' pooled debe estar entre 0 y 1
                    
                    'Se calcula el SE que va a ser el denominador del estadístico
                    se = Sqr((pooled * (1 - pooled)) * ((n2 + n1) / (n2 * n1)))
                    'Calcula la diferencia estandarizada (z-score)
                    If se <> 0 Then
                        valorDif = (p2 - p1) / se
                    Else
                        valorDif = 0 ' Evita errores si el error estándar es cero
                    End If

                    ' Aplica el formato condicional basado en el valor de la diferencia
                    With ws.Cells(fila, ultimaColumna).Font
                        If valorDif >= 1.96 Then
                            .Color = RGB(0, 128, 0) ' Verde (texto)
                            .Bold = True
                        ElseIf valorDif <= -1.96 Then
                            .Color = RGB(255, 0, 0) ' Rojo (texto)
                            .Bold = True
                        Else
                            .Color = RGB(0, 0, 0) ' Negro (texto normal)
                            .Bold = False
                        End If
                    End With
                Else
                End If
            Else
            End If
        Else
        End If
    Next fila
End Sub

