Option Explicit

Const TARGET_URL As String = "file:///C:/Users/juliana%20fabrica/Desktop/listado_clientes.XLS"
Const TARGET_PATH As String = "C:\Users\juliana fabrica\Desktop\listado_clientes.XLS"

' ============================================================
' EVENTO DE APERTURA
'
' No procesa el documento inmediatamente.
' Programa el procesamiento para que Calc pueda terminar
' de entrar en su ciclo normal de interfaz.
' ============================================================

Sub MemoryFix_OnOpen
    Dim oDoc As Object
    Dim oAsyncCallback As Object
    Dim oCallback As Object

    On Error GoTo FatalError

    oDoc = ThisComponent

    If IsNull(oDoc) Then Exit Sub
    If oDoc.URL = "" Then Exit Sub

    ' Ejecutar solamente sobre el XLS generado por G2000.
    If LCase(oDoc.URL) <> LCase(ConvertToURL(TARGET_PATH)) Then Exit Sub

    ' No modificar documentos de solo lectura.
    If oDoc.isReadOnly Then Exit Sub

    ' Crear callback asíncrono.
    oAsyncCallback = createUnoService("com.sun.star.awt.AsyncCallback")

    ' Crear listener para recibir la llamada posterior.
    oCallback = createUnoListener( _
        "MemoryFixAsync_", _
        "com.sun.star.awt.XCallback")

    ' Pasar el documento original al callback.
    oAsyncCallback.addCallback(oCallback, oDoc)

    Exit Sub

FatalError:
    ' Intencionalmente silencioso.
    ' Un error del macro nunca debe impedir que Calc abra el documento.
    Exit Sub
End Sub


' ============================================================
' CALLBACK ASÍNCRONO
'
' Esta rutina se ejecuta posteriormente, cuando OpenOffice
' vuelve a procesar su cola de mensajes.
' ============================================================

Sub MemoryFixAsync_notify(oDoc As Object)
    On Error GoTo FatalError

    If IsNull(oDoc) Then Exit Sub

    ' Aplicar exactamente la lógica original.
    MemoryFix_Apply oDoc

    Exit Sub

FatalError:
    ' Intencionalmente silencioso.
    Exit Sub
End Sub


' ============================================================
' PROCESAMIENTO PRINCIPAL
'
' Esta es esencialmente la rutina original de MemoryFix,
' separada del evento para poder ejecutarla mediante
' AsyncCallback.
' ============================================================

Sub MemoryFix_Apply(oDoc As Object)

    Dim oSheet As Object
    Dim headerRow As Long
    Dim totalRow As Long
    Dim compCol As Long, serieCol As Long, debeCol As Long, haberCol As Long
    Dim lastRow As Long
    Dim r As Long, c As Long
    Dim comp As String, serie As String
    Dim oldVal As Double, newVal As Double
    Dim deltaDebe As Double, deltaHaber As Double
    Dim changedRows As Long

    On Error GoTo FatalError

    If IsNull(oDoc) Then Exit Sub
    If oDoc.URL = "" Then Exit Sub

    ' Seguridad adicional:
    ' volver a comprobar que se trata del XLS de G2000.
    If LCase(oDoc.URL) <> LCase(ConvertToURL(TARGET_PATH)) Then Exit Sub

    If oDoc.isReadOnly Then Exit Sub

    If oDoc.Sheets.getCount() = 0 Then Exit Sub

    oSheet = oDoc.Sheets.getByIndex(0)

    ' --------------------------------------------------------
    ' Buscar fila de encabezados.
    ' --------------------------------------------------------

    headerRow = FindHeaderRow(oSheet)

    If headerRow < 0 Then Exit Sub

    ' --------------------------------------------------------
    ' Buscar columnas.
    ' --------------------------------------------------------

    compCol = FindHeaderColumn(oSheet, headerRow, "COMPROBANTE")
    serieCol = FindHeaderColumn(oSheet, headerRow, "SERIE")
    debeCol = FindHeaderColumn(oSheet, headerRow, "DEBE")
    haberCol = FindHeaderColumn(oSheet, headerRow, "HABER")

    If compCol < 0 Or serieCol < 0 Or debeCol < 0 Or haberCol < 0 Then Exit Sub

    ' --------------------------------------------------------
    ' Obtener última fila utilizada.
    ' --------------------------------------------------------

    lastRow = GetLastUsedRow(oSheet)
    totalRow = -1

    ' --------------------------------------------------------
    ' Buscar primero la fila TOTAL.
    ' --------------------------------------------------------

    For r = headerRow + 1 To lastRow
        If UCase(Trim(oSheet.getCellByPosition(compCol, r).String)) = "TOTAL" Then
            totalRow = r
            Exit For
        End If
    Next r

    ' --------------------------------------------------------
    ' Corregir únicamente documentos conocidos como
    ' devoluciones / notas de crédito.
    ' --------------------------------------------------------

    For r = headerRow + 1 To lastRow

        If r <> totalRow Then

            comp = NormalizeText( _
                oSheet.getCellByPosition(compCol, r).String)

            serie = NormalizeText( _
                oSheet.getCellByPosition(serieCol, r).String)

            If IsCorrective(comp, serie) Then

                ' ------------------------------------------------
                ' DEBE
                ' ------------------------------------------------

                oldVal = oSheet.getCellByPosition(debeCol, r).Value

                If oldVal > 0 Then

                    newVal = -Abs(oldVal)

                    oSheet.getCellByPosition(debeCol, r).Value = newVal

                    deltaDebe = deltaDebe + (newVal - oldVal)

                End If

                ' ------------------------------------------------
                ' HABER
                ' ------------------------------------------------

                oldVal = oSheet.getCellByPosition(haberCol, r).Value

                If oldVal > 0 Then

                    newVal = -Abs(oldVal)

                    oSheet.getCellByPosition(haberCol, r).Value = newVal

                    deltaHaber = deltaHaber + (newVal - oldVal)

                End If

                changedRows = changedRows + 1

            End If

        End If

    Next r

    ' --------------------------------------------------------
    ' Ajustar TOTAL.
    '
    ' Memory/G2000 exporta el TOTAL como valor almacenado,
    ' no como fórmula.
    ' --------------------------------------------------------

    If totalRow >= 0 Then

        If deltaDebe <> 0 Then

            oSheet.getCellByPosition(debeCol, totalRow).Value = _
                oSheet.getCellByPosition(debeCol, totalRow).Value + deltaDebe

        End If

        If deltaHaber <> 0 Then

            oSheet.getCellByPosition(haberCol, totalRow).Value = _
                oSheet.getCellByPosition(haberCol, totalRow).Value + deltaHaber

        End If

    End If

    ' --------------------------------------------------------
    ' Recalcular y guardar.
    ' --------------------------------------------------------

    If changedRows > 0 Then

        oDoc.calculateAll()

        oDoc.store()

    End If

    Exit Sub

FatalError:
    ' Intencionalmente silencioso en producción.
    ' Los errores nunca deben impedir que Calc continúe.
    Exit Sub

End Sub


' ============================================================
' DETERMINAR SI UNA FILA ES CORRECTIVA
' ============================================================

Function IsCorrective(comp As String, serie As String) As Boolean

    Dim isReturn As Boolean
    Dim isNC As Boolean

    isReturn = (comp = "DEV.CONTAD")
    isNC = (comp = "N/CRED")

    If (serie = "NCTK" Or serie = "NCFC") Then

        IsCorrective = (isReturn Or isNC)

    Else

        IsCorrective = False

    End If

End Function


' ============================================================
' NORMALIZAR TEXTO
' ============================================================

Function NormalizeText(s As String) As String

    NormalizeText = UCase(Trim(s))

End Function


' ============================================================
' BUSCAR FILA DE ENCABEZADOS
' ============================================================

Function FindHeaderRow(oSheet As Object) As Long

    Dim r As Long
    Dim hasComp As Boolean
    Dim hasSerie As Boolean
    Dim hasDebe As Boolean
    Dim hasHaber As Boolean

    FindHeaderRow = -1

    For r = 0 To 40

        hasComp = ( _
            FindHeaderColumn(oSheet, r, "COMPROBANTE") >= 0)

        hasSerie = ( _
            FindHeaderColumn(oSheet, r, "SERIE") >= 0)

        hasDebe = ( _
            FindHeaderColumn(oSheet, r, "DEBE") >= 0)

        hasHaber = ( _
            FindHeaderColumn(oSheet, r, "HABER") >= 0)

        If hasComp And hasSerie And hasDebe And hasHaber Then

            FindHeaderRow = r
            Exit Function

        End If

    Next r

End Function


' ============================================================
' BUSCAR COLUMNA POR NOMBRE
' ============================================================

Function FindHeaderColumn( _
    oSheet As Object, _
    rowIndex As Long, _
    wanted As String) As Long

    Dim c As Long
    Dim s As String

    FindHeaderColumn = -1

    For c = 0 To 20

        s = UCase( _
            Trim(oSheet.getCellByPosition(c, rowIndex).String))

        If s = UCase(wanted) Then

            FindHeaderColumn = c
            Exit Function

        End If

    Next c

End Function


' ============================================================
' OBTENER ÚLTIMA FILA UTILIZADA
' ============================================================

Function GetLastUsedRow(oSheet As Object) As Long

    Dim oCursor As Object
    Dim aRange As Object

    oCursor = oSheet.createCursor()

    oCursor.gotoEndOfUsedArea(True)

    aRange = oCursor.getRangeAddress()

    GetLastUsedRow = aRange.EndRow

End Function