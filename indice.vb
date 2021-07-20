Sub crearIndice()

'**************************************
'PASO 1: Crear o limpiar la hoja Indice
'**************************************
Dim hoja As Worksheet
On Error Resume Next
Set hoja = Worksheets("Indice")
On Error GoTo 0

If hoja Is Nothing Then
    'La hoja Indice no existe - Crearla en primera posición
    Worksheets.Add(Before:=Worksheets(1)).Name = "Indice"
Else
    'La hoja Indice ya existe - Limpiarla
    Worksheets("Indice").Cells.Clear
End If

'Insertar título a la hoja Indice
Worksheets("Indice").Range("A1").Value = "Indice"


'************************************************
'PASO 2: Recorrer las hojas creando hipervinculos
'************************************************
Dim fila As Long
Dim vinculoRegreso As String

fila = 2
'Celda donde se colocará el hipervinculo de regreso al indice
vinculoRegreso = "L1"

For Each hoja In Worksheets
    If hoja.Name <> "Indice" Then
        'Crear hipervinculo en hoja Indice
        With Worksheets("Indice")
            .Hyperlinks.Add Anchor:=.Cells(fila, 1), _
            Address:="", _
            SubAddress:="'" & hoja.Name & "'!A1", _
            TextToDisplay:=hoja.Name
        End With
        
        'Crear hipervinculo en hoja destino hacia Indice
        With hoja
            .Hyperlinks.Add Anchor:=.Range(vinculoRegreso), _
            Address:="", _
            SubAddress:="Indice!A1", _
            TextToDisplay:="Indice"
        End With
        fila = fila + 1
    End If
Next

End Sub
