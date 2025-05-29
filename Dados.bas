Attribute VB_Name = "Dados"
' V 1.4.0

Option Explicit

Public SapGuiAuto As Object
Public SAPApplication As Object
Public Connection As Object
Public session As Object
Public colDict As Object ' Dictionary to store column names and indexes

Sub AtualizarDados(Optional ShowOnMacroList As Boolean = False)
        
    On Error GoTo ErrorHandler
    
    OptimizeCodeExecution True
    
    Dim wbThis As Workbook, exportWb As Workbook, wb As Workbook
    Dim wsSource As Worksheet, wsTarget As Worksheet
    Dim sourceLastRow As Long, sourceLastCol As Long, targetLastRow As Long, targetLastCol As Long
    Dim sourceHeaderRow As Range
    Dim sourceColIndex As Long, targetColIndex As Long
    Dim sourcePEP As String, sourceMaterial As String, sourceValor As String, sourceIncoterms As String
    Dim targetPEP As String, targetZETO As String, targetZVA1 As String
    Dim i As Long, j As Long
    Dim cell As Range
    Dim exportWbName As String, exportWbPath As String
    Dim tries As Long
    Dim found As Boolean, isNotFound As Boolean
    Dim ErrSection As String
    Dim sourceColDict As Object
    Dim startTime As Double
    Dim key As Variant
    Dim amount As Double
    
ErrSection = "variableDeclarations"
    
    Set wbThis = ThisWorkbook
    
    Set wsTarget = wbThis.Sheets("FATURAMENTO")
    
ErrSection = "variableDeclarations10"

    If wsTarget Is Nothing Then
        GoTo ErrorHandler
    End If
    
    ' Clear all filters if any
    If wsTarget.ListObjects(1).ShowAutoFilter Then
        wsTarget.ListObjects(1).AutoFilter.ShowAllData
    End If
    
    Call GetAllColumnIndexes(wsTarget)
    
    For Each key In colDict.Keys
        If colDict(key) = 0 Then
            MsgBox "Uma coluna não foi encontrada." & vbCrLf & key & vbCrLf & "A macro será encerrada.", vbExclamation, "Falha ao Mapear Colunas"
            GoTo CleanExit
        End If
    Next key
    
    ' Setup SAP and check if it is running
    Do While Not SetupSAPScripting
        ' Ask the user to initiate SAP or cancel
        Dim response As VbMsgBoxResult
        response = MsgBox("SAP não está acessível. Inicie o SAP e clique em OK para tentar novamente, ou Cancelar para sair.", vbOKCancel + vbExclamation, "Aguardando SAP")
    
        If response = vbCancel Then
            MsgBox "Execução terminada pelo usuário.", vbInformation
            GoTo CleanExit  ' Exit the function or sub
        End If
    Loop
    
ErrSection = "extractZTMM091FromSAP"

    ' Name of the workbook to find
    exportWbName = "ZTMM091"
    
    ' SAP Navigation and Export
    session.findById("wnd[0]/tbar[0]/okcd").Text = "/nZTMM091"
    session.findById("wnd[0]").sendVKey 0
    
    session.findById("wnd[0]/tbar[1]/btn[17]").press
    session.findById("wnd[1]/usr/txtENAME-LOW").Text = "JULIANARIGO"
    session.findById("wnd[1]/tbar[0]/btn[8]").press
    session.findById("wnd[0]/tbar[1]/btn[8]").press
    session.findById("wnd[0]/tbar[1]/btn[43]").press

    ' Close the file extension pop-up
    On Error Resume Next
    If session.findById("wnd[1]/usr/ctxtDY_FILENAME") Is Nothing Then
        session.findById("wnd[1]/tbar[0]/btn[0]").press
    End If
    On Error GoTo ErrorHandler
    
    exportWbName = Replace(session.findById("wnd[1]/usr/ctxtDY_FILENAME").Text, "export", exportWbName)
    exportWbPath = session.findById("wnd[1]/usr/ctxtDY_PATH").Text
    
    session.findById("wnd[1]/usr/ctxtDY_FILENAME").Text = exportWbName
    session.findById("wnd[1]/tbar[0]/btn[11]").press

ErrSection = "extractDataFromZTMM091"
    
    startTime = Timer
    
    Do
        found = False
        
        ' Loop through all open workbooks
        For Each wb In Application.Workbooks
            If UCase(wb.Name) = UCase(exportWbName) Then
                Set exportWb = wb
                found = True
                Exit Do  ' Exit the loop immediately if the workbook is found
            End If
        Next wb
        
        ' Check if 60 seconds have elapsed
        If Timer - startTime >= 120 Then
            ErrSection = "extractDataFromZTMM091-timeout"
            MsgBox "Erro de Timeout. Não foi possível atualizar conforme VL10G.", vbInformation, "Timeout"
            GoTo extractVL10GFromSAP
        End If
        
        DoEvents  ' Yield control to allow other events to be processed
        
    Loop

    ' Set worksheets
    Set wsSource = exportWb.Sheets(1)
    Set wsTarget = wbThis.Sheets("FATURAMENTO")

ErrSection = "extractDataFromZTMM09110"

    If wsSource Is Nothing Or wsTarget Is Nothing Then
        GoTo ErrorHandler
    End If

    ' Find the last used row and column in the source sheet
    targetLastRow = wsTarget.Cells(wsSource.Rows.Count, 1).End(xlUp).Row
    targetLastCol = wsTarget.Cells(1, wsSource.Columns.Count).End(xlToLeft).Column

ErrSection = "extractDataFromZTMM09120"

    ' If "Situação" column not found, exit sub
    If colDict("Status") = 0 Then
        GoTo ErrorHandler
    End If
    
    ' Find the last used row and column in the source sheet
    sourceLastRow = wsSource.Cells(wsSource.Rows.Count, 1).End(xlUp).Row
    sourceLastCol = wsSource.Cells(1, wsSource.Columns.Count).End(xlToLeft).Column

    ' Find the "Situação" column
    Set sourceHeaderRow = wsSource.Range(wsSource.Cells(1, 1), wsSource.Cells(1, sourceLastCol))
    sourceColIndex = 0
    
    For Each cell In sourceHeaderRow
        If Trim(UCase(cell.Value)) = "MATERIAL" Then
            sourceColIndex = cell.Column
            Exit For
        End If
    Next cell
    
ErrSection = "extractDataFromZTMM09130"

    ' If column not found, exit sub
    If sourceColIndex = 0 Then
        GoTo ErrorHandler
    End If
    
    ' Loop through rows from bottom to top to avoid skipping rows after deletion
    For i = sourceLastRow To 2 Step -1 ' Assuming headers are in row 1
ErrSection = "extractDataFromZTMM09140-" & i
        isNotFound = True
        
        sourcePEP = wsSource.Cells(i, sourceColIndex + 2).Value
        sourceMaterial = wsSource.Cells(i, sourceColIndex).Value
        sourceValor = wsSource.Cells(i, sourceColIndex + 8).Value
        
        If sourcePEP = "" Then
            GoTo SkipIterationZTMM091
        End If
        
        For j = targetLastRow To 2 Step -1 ' Assuming headers are in row 1
            targetPEP = wsTarget.Cells(j, colDict("PEP")).Value
            targetZETO = wsTarget.Cells(j, colDict("ZETO")).Value
            targetZVA1 = wsTarget.Cells(j, colDict("ZVA1")).Value
        
            If sourcePEP = targetPEP And (sourceMaterial = targetZETO Or sourceMaterial = targetZVA1) Then
                isNotFound = False
                Exit For
            End If
        Next j
        
        If isNotFound Then
            ' Create a new row with OrderLocation "Itajaí" and PhysicalStock 1321 (without Incoterm)
            Call AddNewRow(wsTarget, colDict, Date, sourcePEP, sourceMaterial, "ITJ", 1321, "")
        Else
            ' Update the existing row at index j with the same values
            Call UpdateRowIfEmpty(wsTarget, j, colDict, Date, sourcePEP, "", "", "", sourceMaterial, "ITJ", "", "", "", "", 1321, "", "")
        End If
SkipIterationZTMM091:
    Next i
    
    exportWb.Close False
    
    On Error Resume Next
    Kill exportWbPath & exportWbName
    On Error GoTo ErrorHandler

extractVL10GFromSAP:
ErrSection = "extractVL10GFromSAP"

    ' Name of the workbook to find
    exportWbName = "VL10G"
    
    ' SAP Navigation and Export
    session.findById("wnd[0]/tbar[0]/okcd").Text = "/nVL10G"
    session.findById("wnd[0]").sendVKey 0
    
    session.findById("wnd[0]/tbar[1]/btn[17]").press
    session.findById("wnd[1]/usr/txtENAME-LOW").Text = "JULIANARIGO"
    session.findById("wnd[1]/tbar[0]/btn[8]").press
   
    Dim grid As Object
    Dim iRow As Long
    Dim searchValue As String
    Dim colName As String
    
    ' Set your search value and the column name (as defined in the grid)
    searchValue = "ENGE JULI"
    colName = "VARIANT"   ' Replace with the actual column name
    
    ' Get the grid control
    Set grid = session.findById("wnd[1]/usr/cntlALV_CONTAINER_1/shellcont/shell")
    
    ' Loop through all rows
    For iRow = 0 To grid.RowCount - 1
        If grid.GetCellValue(iRow, colName) = searchValue Then
            ' When found, set the current cell to the matching row
            grid.CurrentCellRow = iRow
            ' Depending on your setup, SelectedRows may require a string
            grid.SelectedRows = CStr(iRow)
            ' Double-click the cell to perform the action
            grid.doubleClickCurrentCell
            Exit For   ' Exit the loop once the desired row is found
        End If
    Next iRow
    
    session.findById("wnd[0]/tbar[1]/btn[8]").press
    session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell").contextMenu
    session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell").selectContextMenuItem "&XXL"
    
    ' Close the file extension pop-up
    On Error Resume Next
    If session.findById("wnd[1]/usr/ctxtDY_FILENAME") Is Nothing Then
        session.findById("wnd[1]/tbar[0]/btn[0]").press
    End If
    On Error GoTo ErrorHandler
    
    exportWbName = Replace(session.findById("wnd[1]/usr/ctxtDY_FILENAME").Text, "export", exportWbName)
    exportWbPath = session.findById("wnd[1]/usr/ctxtDY_PATH").Text
    
    session.findById("wnd[1]/usr/ctxtDY_FILENAME").Text = exportWbName
    session.findById("wnd[1]/tbar[0]/btn[11]").press

ErrSection = "extractDataFromVL10G"

    startTime = Timer
    
    Do
        found = False
        
        ' Loop through all open workbooks
        For Each wb In Application.Workbooks
            If UCase(wb.Name) = UCase(exportWbName) Then
                Set exportWb = wb
                found = True
                Exit Do  ' Exit the loop immediately if the workbook is found
            End If
        Next wb
        
        ' Check if 60 seconds have elapsed
        If Timer - startTime >= 120 Then
            ErrSection = "extractDataFromVL10G-timeout"
            MsgBox "Erro de Timeout. Não foi possível atualizar conforme VL10G.", vbInformation, "Timeout"
            GoTo completeInformationFromAnalisys
        End If
        
        DoEvents  ' Yield control to allow other events to be processed
        
        ' Pause for 3 seconds to give the external process time to open the workbook
        Application.Wait Now + TimeValue("00:00:03")
    Loop

    ' Set worksheets
    Set wsSource = exportWb.Sheets(1)
    Set wsTarget = wbThis.Sheets("FATURAMENTO")

ErrSection = "extractDataFromVL10G10"
        
     If wsSource Is Nothing Or wsTarget Is Nothing Then
        GoTo ErrorHandler
    End If

    ' Find the last used row and column in the source sheet
    targetLastRow = wsTarget.Cells(wsSource.Rows.Count, 1).End(xlUp).Row
    targetLastCol = wsTarget.Cells(1, wsSource.Columns.Count).End(xlToLeft).Column
    
ErrSection = "extractDataFromVL10G20"

    ' If "Situação" column not found, exit sub
    If colDict("Status") = 0 Then
        GoTo ErrorHandler
    End If
    
    ' Find the last used row and column in the source sheet
    sourceLastRow = wsSource.Cells(wsSource.Rows.Count, 1).End(xlUp).Row
    sourceLastCol = wsSource.Cells(1, wsSource.Columns.Count).End(xlToLeft).Column

    ' Find the "Situação" column
    Set sourceHeaderRow = wsSource.Range(wsSource.Cells(1, 1), wsSource.Cells(1, sourceLastCol))
    sourceColIndex = 0
    
    For Each cell In sourceHeaderRow
        If Trim(UCase(cell.Value)) = "MATERIAL" Then
            sourceColIndex = cell.Column
            Exit For
        End If
    Next cell

ErrSection = "extractDataFromVL10G30"

    ' If column not found, exit sub
    If sourceColIndex = 0 Then
        GoTo ErrorHandler
    End If
    
    ' Loop through rows from bottom to top to avoid skipping rows after deletion
    For i = sourceLastRow To 2 Step -1 ' Assuming headers are in row 1
ErrSection = "extractDataFromVL10G40-" & i
        isNotFound = True
        
        sourceIncoterms = wsSource.Cells(i, sourceColIndex - 7).Value
        sourcePEP = wsSource.Cells(i, sourceColIndex - 3).Value
        sourceMaterial = wsSource.Cells(i, sourceColIndex).Value
        sourceValor = wsSource.Cells(i, sourceColIndex + 7).Value
        
        If sourcePEP = "" Then
            GoTo SkipIterationVL10G
        End If
        
        For j = targetLastRow To 2 Step -1 ' Assuming headers are in row 1
            targetPEP = wsTarget.Cells(j, colDict("PEP")).Value
            targetZETO = wsTarget.Cells(j, colDict("ZETO")).Value
            targetZVA1 = wsTarget.Cells(j, colDict("ZVA1")).Value
        
            If sourcePEP = targetPEP And (sourceMaterial = targetZETO Or sourceMaterial = targetZVA1) Then
                isNotFound = False
                Exit For
            End If
        Next j
        
        If isNotFound Then
            ' Create a new row with OrderLocation "Jaraguá", PhysicalStock 1320 and include Incoterm
            Call AddNewRow(wsTarget, colDict, Date, sourcePEP, sourceMaterial, "JGS", 1320, sourceIncoterms)
        Else
            ' Update the existing row at index j
            Call UpdateRowIfEmpty(wsTarget, j, colDict, Date, sourcePEP, "", "", "", sourceMaterial, "JGS", sourceIncoterms, "", "", "", 1321, "", "")
        End If
SkipIterationVL10G:
    Next i
    
    exportWb.Close False
    
    On Error Resume Next
    Kill exportWbPath & exportWbName
    On Error GoTo ErrorHandler

completeInformationFromAnalisys:
ErrSection = "completeInformationFromAnalisys"

    ' Open the workbook Reunião de Faturamento Semanal - New Layout.xlsm
    Workbooks.Open "\\brjgs100\DFSWEG\GROUPS\BR_SC_JGS_WAU_ADM_CONTRATOS\ACIONAMENTOS\00-EQUIPE DE APOIO\00-BANCO DE DADOS\ANALYSIS_ADCON_WAU.xlsm"
    exportWbName = "ANALYSIS_ADCON_WAU.xlsm"
    
    tries = 0

    Do
        If tries > 10 Then
            GoTo ErrorHandler
        End If
        
        found = False
        
        ' Loop through all open workbooks
        For Each wb In Application.Workbooks
            If UCase(wb.Name) = UCase(exportWbName) Then
                Set exportWb = wb
                found = True
                Exit Do
            End If
        Next wb
    
        tries = tries + 1
        
        DoEvents
    Loop

    ' Set worksheets
    Set wsSource = exportWb.Sheets("ADCON_WAU GERAL")
    Set wsTarget = wbThis.Sheets("FATURAMENTO")

ErrSection = "completeInformationFromAnalisys10"

     If wsSource Is Nothing Or wsTarget Is Nothing Then
        GoTo ErrorHandler
    End If

    ' Find the last used row and column in the source sheet
    targetLastRow = wsTarget.Cells(wsSource.Rows.Count, 1).End(xlUp).Row
    targetLastCol = wsTarget.Cells(2, wsSource.Columns.Count).End(xlToLeft).Column

ErrSection = "completeInformationFromAnalisys20"

    ' If "Situação" column not found, exit sub
    If colDict("Status") = 0 Then
        GoTo ErrorHandler
    End If
    
    ' Find the last used row and column in the source sheet
    sourceLastRow = wsSource.Cells(wsSource.Rows.Count, 2).End(xlUp).Row
    sourceLastCol = wsSource.Cells(2, wsSource.Columns.Count).End(xlToLeft).Column

ErrSection = "completeInformationFromAnalisys30"

    ' Find the source columns
    Set sourceColDict = CreateObject("Scripting.Dictionary")
    Set sourceColDict = GetSourceColumnIndexes(wsSource)
    
    For Each key In sourceColDict.Keys
        If sourceColDict(key) = 0 Then
            MsgBox "Uma coluna do Analysis não foi encontrada." & vbCrLf & key & vbCrLf & "Os dados não serão atualizados do Analysis.", vbExclamation, "Falha ao Mapear Colunas"
            GoTo completeLocationInfo
        End If
    Next key
    
    ' Loop through rows from bottom to top to avoid skipping rows after deletion
    For i = sourceLastRow To 2 Step -1 ' Assuming headers are in row 1
ErrSection = "completeInformationFromAnalisys40-" & i
        isNotFound = True

        If InStr(wsSource.Cells(i, sourceColDict("PEP")).Value, "-") <> 0 Then
            For j = targetLastRow To 3 Step -1 ' Assuming headers are in row 1
                If Left(wsSource.Cells(i, sourceColDict("PEP")).Value, InStr(InStr(wsSource.Cells(i, sourceColDict("PEP")).Value, "-") + 1, wsSource.Cells(i, sourceColDict("PEP")).Value, "-") - 1) = Left(wsTarget.Cells(j, colDict("PEP")).Value, InStr(InStr(wsTarget.Cells(j, colDict("PEP")).Value, "-") + 1, wsTarget.Cells(j, colDict("PEP")).Value, "-") - 1) Then
                    isNotFound = False
                    Exit For
                End If
            Next j
        End If
        
        If isNotFound Then
            ' Create a new row with OrderLocation "Jaraguá", PhysicalStock 1320 and include Incoterm
            ' Call AddNewRow(wsTarget, colDict, Date, sourcePEP, sourceMaterial, "Jaraguá", 1320, sourceIncoterms)
        Else
            
            ' Check in wich column is the correct value for the following call
            If wsSource.Cells(i, sourceColDict("Wallet")).Value > wsSource.Cells(i, sourceColDict("Amount")).Value Then
                amount = wsSource.Cells(i, sourceColDict("Wallet")).Value
            Else
                amount = wsSource.Cells(i, sourceColDict("Amount")).Value
            End If
        
            ' Update the existing row at index j
            Call UpdateRowIfEmpty(wsTarget, j, colDict, Date, wsSource.Cells(i, sourceColDict("PEP")).Value, wsSource.Cells(i, sourceColDict("Market")).Value, wsSource.Cells(i, sourceColDict("Client")).Value, wsSource.Cells(i, sourceColDict("SalesDoc")).Value, "", "", wsSource.Cells(i, sourceColDict("Incoterms")).Value, wsSource.Cells(i, sourceColDict("Incoterms2")).Value, wsSource.Cells(i, sourceColDict("PM")).Value, amount, wsSource.Cells(i, sourceColDict("Plant")).Value, wsSource.Cells(i, sourceColDict("PrepDate")).Value, wsSource.Cells(i, sourceColDict("ShipmentDate")).Value)
        End If
    Next i

completeLocationInfo:
ErrSection = "completeLocationInfo"
    
    Set wsTarget = wbThis.Sheets("FATURAMENTO")
    
     If wsTarget Is Nothing Then
        GoTo ErrorHandler
    End If

ErrSection = "completeLocationInfo10"

    ' Find the last used row and column in the source sheet
    targetLastRow = wsTarget.Cells(wsSource.Rows.Count, 1).End(xlUp).Row
    targetLastCol = wsTarget.Cells(1, wsSource.Columns.Count).End(xlToLeft).Column

ErrSection = "completeLocationInfo20"

    ' If "Situação" column not found, exit sub
    If colDict("Status") = 0 Then
        GoTo ErrorHandler
    End If
    
    ' Loop through rows from bottom to top to avoid skipping rows after deletion
    For i = targetLastRow To 2 Step -1 ' Assuming headers are in row 1
ErrSection = "completeLocationInfo30-" & i
        With wsTarget.Cells(i, colDict("PhysicalStock"))
            If wsTarget.Cells(i, colDict("OrderLocation")).Value = "" Then
                ' Trim to first 4 characters if longer than 4
                If Len(.Value) > 4 Then
                    .Value = Left(.Value, 4)
                End If
                
                ' Assign OrderLocation based on PhysicalStock
                Select Case .Value
                    Case "1320"
                        wsTarget.Cells(i, colDict("OrderLocation")).Value = "JGS"
                    Case "1321"
                        wsTarget.Cells(i, colDict("OrderLocation")).Value = "ITJ"
                End Select
            End If
            
            If wsTarget.Cells(i, colDict("PhysicalStock")).Value = "" Then
                ' Assign PhysicalStock based on OrderLocation
                If InStr(1, UCase(wsTarget.Cells(i, colDict("OrderLocation")).Value), "JGS") > 0 Then
                    wsTarget.Cells(i, colDict("PhysicalStock")).Value = 1320
                ElseIf InStr(1, UCase(wsTarget.Cells(i, colDict("OrderLocation")).Value), "ITJ") > 0 Then
                    wsTarget.Cells(i, colDict("PhysicalStock")).Value = 1321
                End If
            End If
        End With
    Next i
    
    ' Success message
    MsgBox "Os dados foram atualizados com sucesso!", vbInformation, "Macro Finalizada"

' Ignore next bit of code
GoTo CleanExit
ErrSection = "moveFinishedItems"

    ' Set worksheets
    Set wsSource = wbThis.Sheets("FATURAMENTO")
    Set wsTarget = wbThis.Sheets("Finalizado")
    
ErrSection = "moveFinishedItems10"

    If wsSource Is Nothing Or wsTarget Is Nothing Then
        GoTo ErrorHandler
    End If

    ' Find the last used row and column in the source sheet
    sourceLastRow = wsSource.Cells(wsSource.Rows.Count, 1).End(xlUp).Row
    sourceLastCol = wsSource.Cells(1, wsSource.Columns.Count).End(xlToLeft).Column

ErrSection = "moveFinishedItems20"

    ' If "Situação" column not found, exit sub
    If colDict("Status") = 0 Then
        GoTo ErrorHandler
    End If
    
    ' Loop through rows from bottom to top to avoid skipping rows after deletion
    For i = sourceLastRow To 2 Step -1 ' Assuming headers are in row 1
ErrSection = "moveFinishedItems30-" & i
        If UCase(Trim(wsSource.Cells(i, colDict("Status")).Value)) = "RECONHECIDO" Then
            ' Find last row in target sheet
            targetLastRow = wsTarget.Cells(wsTarget.Rows.Count, 1).End(xlUp).Row + 1

            ' Copy and paste formats from the row above
            wsTarget.Rows(targetLastRow - 1).Copy
            wsTarget.Rows(targetLastRow).PasteSpecial Paste:=xlFormats
            
            ' Copy and paste values from the source
            wsSource.Rows(i).Copy
            wsTarget.Rows(targetLastRow).PasteSpecial Paste:=xlValue
            
            ' Clear clipboard
            Application.CutCopyMode = False
            
            ' Delete the original row to avoid empty rows
            wsSource.Rows(i).Delete Shift:=xlUp
        End If
    Next i

    ' Clear clipboard
    Application.CutCopyMode = False

    Dim shp As Shape
    Dim txtBox As Shape
    Dim shapeFound As Boolean
    
    Set wsSource = wbThis.Sheets("FATURAMENTO")

    ' Loop through all shapes in the sheet
    shapeFound = False
    For Each shp In wsSource.Shapes
        ' Check if the shape is a text box
        If shp.Type = msoTextBox Then
            shp.TextFrame2.TextRange.Text = "Última atualização: " & vbCrLf & Now
            shapeFound = True
            Exit For
        End If
    Next shp

CleanExit:

    ' Loop through all open workbooks to find if the exportWb is oppened
    For Each wb In Application.Workbooks
        If UCase(wb.Name) = UCase(exportWbName) Then
            wb.Close False ' Close the exportWb
        End If
    Next wb
    
    ' Ensure that all optimizations are turned back on
    OptimizeCodeExecution False
    Exit Sub

ErrorHandler:
        
    ' Log and diagnose the error using Erl to show the last executed line number
    MsgBox "Erro " & Err.Number & " após " & ErrSection & ": " & Err.Description, vbCritical, "Erro em AtualizarDados"
    
    ' Resume cleanup to ensure that settings are restored
    GoTo CleanExit
End Sub

Sub AddNewRow(wsTarget As Worksheet, colDict As Object, _
                   sourceDate As Date, sourcePEP As Variant, sourceMaterial As Variant, _
                   orderLocation As String, physicalStock As Variant, _
                   sourceIncoterms As Variant)
    Dim tbl As ListObject
    Dim newRow As ListRow
    ' Reference the table in wsTarget
    Set tbl = wsTarget.ListObjects("Tabela1")
    ' Add a new row to the table
    Set newRow = tbl.ListRows.Add
    
    ' Populate the new row with values
    newRow.Range.Cells(1, colDict("Date")).Value = sourceDate
    newRow.Range.Cells(1, colDict("PEP")).Value = sourcePEP
    newRow.Range.Cells(1, colDict("ZETO")).Value = sourceMaterial
    newRow.Range.Cells(1, colDict("ZVA1")).Value = sourceMaterial
    newRow.Range.Cells(1, colDict("OrderLocation")).Value = orderLocation
    newRow.Range.Cells(1, colDict("Incoterm")).Value = sourceIncoterms
    newRow.Range.Cells(1, colDict("StockStatus")).Value = "OK"
    newRow.Range.Cells(1, colDict("Checklist")).Value = "PENDENTE"
    newRow.Range.Cells(1, colDict("Freight")).Value = "PENDENTE"
    newRow.Range.Cells(1, colDict("Status")).Value = "AGUARD. PM"
    newRow.Range.Cells(1, colDict("PhysicalStock")).Value = physicalStock
    
End Sub

Sub UpdateRowIfEmpty(wsTarget As Worksheet, rowIndex As Long, colDict As Object, _
                     sourceDate As Date, sourcePEP As Variant, sourceMarket As Variant, sourceClient As Variant, _
                     sourceOV As Variant, sourceMaterial As Variant, sourceOrderLocation As String, _
                     sourceIncoterms As Variant, sourceIncoterms2 As Variant, sourcePM As Variant, _
                     sourceAmount As Variant, sourcePhysicalStock As Variant, sourcePrepDate As Variant, sourceShipmentDate As Variant)
                
    With wsTarget
    
        ' Update Date if empty, Column A
        If False Then
            .Cells(rowIndex, colDict("Date")).Value = sourceDate
        End If
        
        ' Update PEP if empty, Column B
        If .Cells(rowIndex, colDict("PEP")).Value = "" Then
            .Cells(rowIndex, colDict("PEP")).Value = sourcePEP
        End If
        
        ' Update Market if empty, Column C
        If .Cells(rowIndex, colDict("Market")).Value = "" Then
            If InStr(1, UCase(sourceMarket), "FORA") > 0 Then
                .Cells(rowIndex, colDict("Market")).Value = "EXTERNO"
            Else
                .Cells(rowIndex, colDict("Market")).Value = "INTERNO"
            End If
        End If
        
        ' Update Client if empty, Column D
        If .Cells(rowIndex, colDict("Client")).Value = "" Then
            .Cells(rowIndex, colDict("Client")).Value = sourceClient
        End If
        
        ' Update OV if empty, Column E
        If .Cells(rowIndex, colDict("OV")).Value = "" Then
            .Cells(rowIndex, colDict("OV")).Value = sourceOV
        End If
        
        ' For ZETO (and ZVA1) update if both are empty, Column G
        If .Cells(rowIndex, colDict("ZETO")).Value = "" And .Cells(rowIndex, colDict("ZVA1")).Value = "" Then
            If .Cells(rowIndex, colDict("Market")).Value = "INTERNO" Then
                .Cells(rowIndex, colDict("ZETO")).Value = ""
                .Cells(rowIndex, colDict("ZVA1")).Value = sourceMaterial
            ElseIf .Cells(rowIndex, colDict("Market")).Value = "EXTERNO" Then
                .Cells(rowIndex, colDict("ZETO")).Value = sourceMaterial
                .Cells(rowIndex, colDict("ZVA1")).Value = ""
            End If
        End If
        
        ' Update OrderLocation if empty, Column I
        If .Cells(rowIndex, colDict("OrderLocation")).Value = "" Then
            .Cells(rowIndex, colDict("OrderLocation")).Value = sourceOrderLocation
        End If
        
        ' Update Incoterm if provided and the cell is empty, Column J
        If .Cells(rowIndex, colDict("Incoterm")).Value = "" Then
            .Cells(rowIndex, colDict("Incoterm")).Value = sourceIncoterms
        End If
        
        ' Update Incoterm2 if provided and the cell is empty, Column K
        If .Cells(rowIndex, colDict("Incoterm2")).Value = "" Then
            .Cells(rowIndex, colDict("Incoterm2")).Value = sourceIncoterms2
        End If
        
        ' Update PM if provided and the cell is empty, Column L
        If .Cells(rowIndex, colDict("PM")).Value = "" Then
            .Cells(rowIndex, colDict("PM")).Value = sourcePM
        End If
    
        ' Update Amount if provided and the cell is empty, Column M
        If .Cells(rowIndex, colDict("Amount")).Value = "" Then
            .Cells(rowIndex, colDict("Amount")).Value = sourceAmount
        End If
        
        ' Update DataReme if empty
        If .Cells(rowIndex, colDict("DataReme")).Value = "" Then
            .Cells(rowIndex, colDict("DataReme")).Value = sourcePrepDate
        End If
        
        ' Update DataPrep if empty
        If .Cells(rowIndex, colDict("DataPrep")).Value = "" Then
            .Cells(rowIndex, colDict("DataPrep")).Value = sourceShipmentDate
        End If
        
        ' Update PhysicalStock if empty
        If .Cells(rowIndex, colDict("PhysicalStock")).Value = "" Then
            .Cells(rowIndex, colDict("PhysicalStock")).Value = sourcePhysicalStock
        End If
    End With
End Sub

Function GetSourceColumnIndexes(ws As Worksheet, Optional ShowOnMacroList As Boolean = False) As Object
    Dim SourceColumnIndexes As Object
    Set SourceColumnIndexes = CreateObject("Scripting.Dictionary")
    
    ' Map your internal aliases to actual header names
    SourceColumnIndexes.Add "Market", GetSourceColumnIndex(ws, "Mercado", 2, SourceColumnIndexes)                     ' Column B
    'SourceColumnIndexes.Add "YearBI", GetSourceColumnIndex(ws, "Ano BI", 2, SourceColumnIndexes)                      ' Column C
    'SourceColumnIndexes.Add "MonthBI", GetSourceColumnIndex(ws, "Mês BI", 2, SourceColumnIndexes)                     ' Column D
    'SourceColumnIndexes.Add "Status", GetSourceColumnIndex(ws, "STATUS", 2, SourceColumnIndexes)                      ' Column E
    SourceColumnIndexes.Add "PurchaseDoc", GetSourceColumnIndex(ws, "Doc. Compra", 2, SourceColumnIndexes)            ' Column F
    SourceColumnIndexes.Add "SalesDoc", GetSourceColumnIndex(ws, "Doc. Vendas", 2, SourceColumnIndexes)               ' Column G
    SourceColumnIndexes.Add "SalesItem", GetSourceColumnIndex(ws, "Item Doc. Venda", 2, SourceColumnIndexes)          ' Column H
    SourceColumnIndexes.Add "PEP", GetSourceColumnIndex(ws, "PEP", 2, SourceColumnIndexes)                            ' Column I
    SourceColumnIndexes.Add "Client", GetSourceColumnIndex(ws, "Cliente", 2, SourceColumnIndexes)                     ' Column J
    SourceColumnIndexes.Add "Incoterms", GetSourceColumnIndex(ws, "Incoterms 1", 2, SourceColumnIndexes)                ' Column K
    SourceColumnIndexes.Add "Incoterms2", GetSourceColumnIndex(ws, "Incoterms 2", 2, SourceColumnIndexes)                ' Column K
    SourceColumnIndexes.Add "PrepDate", GetSourceColumnIndex(ws, "Data Prep. Material", 2, SourceColumnIndexes)       ' Column L
    SourceColumnIndexes.Add "ShipmentDate", GetSourceColumnIndex(ws, "Data Remessa", 2, SourceColumnIndexes)          ' Column M
    'SourceColumnIndexes.Add "DataAdcB", GetSourceColumnIndex(ws, "Data Adc. B", 2, SourceColumnIndexes)               ' Column N
    'SourceColumnIndexes.Add "InvoiceDate", GetSourceColumnIndex(ws, "Data NF", 2, SourceColumnIndexes)                ' Column O
    'SourceColumnIndexes.Add "PCPDate", GetSourceColumnIndex(ws, "Data PCP", 2, SourceColumnIndexes)                   ' Column P
    'SourceColumnIndexes.Add "RevenueReceivedDate", GetSourceColumnIndex(ws, "Data de Rec.Receita", 2, SourceColumnIndexes) ' Column Q
    'SourceColumnIndexes.Add "CauseArea", GetSourceColumnIndex(ws, "Área Causadora", 2, SourceColumnIndexes)           ' Column R
    'SourceColumnIndexes.Add "Reason", GetSourceColumnIndex(ws, "Motivo", 2, SourceColumnIndexes)                      ' Column S
    'SourceColumnIndexes.Add "Notes", GetSourceColumnIndex(ws, "Observação", 2, SourceColumnIndexes)                   ' Column T
    'SourceColumnIndexes.Add "PreviousMeeting", GetSourceColumnIndex(ws, "Reunião Anterior", 2, SourceColumnIndexes)   ' Column U
    'SourceColumnIndexes.Add "Week", GetSourceColumnIndex(ws, "Semana", 2, SourceColumnIndexes)                        ' Column V
    SourceColumnIndexes.Add "PM", GetSourceColumnIndex(ws, "Funcionário Responsável", 2, SourceColumnIndexes)                              ' Column W
    SourceColumnIndexes.Add "Wallet", GetSourceColumnIndex(ws, "Vlr. Carteira", 1, SourceColumnIndexes)                       ' Column X
    SourceColumnIndexes.Add "Amount", GetSourceColumnIndex(ws, "Vlr. ROL", 1, SourceColumnIndexes)
    SourceColumnIndexes.Add "Plant", GetSourceColumnIndex(ws, "Centro", 2, SourceColumnIndexes)                       ' Column Y
    SourceColumnIndexes.Add "ProductHierarchy", GetSourceColumnIndex(ws, "Hier. Produto", 2, SourceColumnIndexes)     ' Column Z
    'SourceColumnIndexes.Add "Month", GetSourceColumnIndex(ws, "Mês", 2, SourceColumnIndexes)                          ' Column AA
    'SourceColumnIndexes.Add "Year", GetSourceColumnIndex(ws, "Ano", 2, SourceColumnIndexes)                           ' Column AB

    Set GetSourceColumnIndexes = SourceColumnIndexes
End Function

Function GetSourceColumnIndex(ws As Worksheet, headerName As String, headerRow As Long, sourceColDict As Object) As Long
    Dim col As Range
    Dim alreadyUsed As Boolean
    
    GetSourceColumnIndex = 0 ' Not found
    
    For Each col In ws.Rows(headerRow).Cells
        If InStr(1, Trim(UCase(col.Value)), Trim(UCase(headerName))) > 0 Then
            alreadyUsed = False

            ' Check if this column number is already assigned in colDict
            Dim key As Variant
            For Each key In sourceColDict.Keys
                If sourceColDict(key) = col.Column Then
                    alreadyUsed = True
                    Exit For
                End If
            Next key

            ' If not already used, assign it
            If Not alreadyUsed Then
                GetSourceColumnIndex = col.Column
                Exit Function
            End If
        End If
    Next col
End Function

Sub GetAllColumnIndexes(ws As Worksheet, Optional ShowOnMacroList As Boolean = False)
    Set colDict = CreateObject("Scripting.Dictionary")
    
    ' Map your internal aliases to actual header names
    colDict.Add "Date", 1               ' Column A
    colDict.Add "PEP", 2                ' Column B
    colDict.Add "Market", 3             ' Column C
    colDict.Add "Client", 4             ' Column D
    colDict.Add "OV", 5                 ' Column E
    colDict.Add "ZVA1", 6               ' Column F
    colDict.Add "ZETO", 7               ' Column G
    colDict.Add "PaymentTerms", 8       ' Column H
    colDict.Add "OrderLocation", 9      ' Column I
    colDict.Add "Incoterm", 10          ' Column J
    colDict.Add "Incoterm2", 11         ' Column K
    colDict.Add "PM", 12                ' Column L
    colDict.Add "Amount", 13            ' Column M
    colDict.Add "DataReme", 14          ' Column N
    colDict.Add "DataPrep", 15          ' Column O
    colDict.Add "BillingResp", 16       ' Column P
    colDict.Add "BillingForecast", 17   ' Column Q
    colDict.Add "StockStatus", 18       ' Column R
    colDict.Add "Checklist", 19         ' Column S
    colDict.Add "Freight", 20           ' Column T
    colDict.Add "Status", 21            ' Column U
    colDict.Add "PhysicalStock", 22     ' Column V
    colDict.Add "STATUS BI", 23         ' Column W
    colDict.Add "ETD", 24               ' Column X
    colDict.Add "ShippingDate", 25      ' Column Y
    colDict.Add "ETA", 26               ' Column Z
    colDict.Add "ShipmentNumber", 27    ' Column AA
    colDict.Add "LogisticsCoord", 28    ' Column AB
    colDict.Add "VALOR BI", 29          ' Column AC
    colDict.Add "Notes", 30             ' Column AD
    
End Sub

Function GetColumnIndex(ws As Worksheet, headerName As String, Optional headerRow As Long = 1) As Long
    Dim col As Range
    Dim alreadyUsed As Boolean
    
    GetColumnIndex = 0 ' Not found
    
    For Each col In ws.Rows(headerRow).Cells
        If InStr(1, Trim(UCase(col.Value)), Trim(UCase(headerName))) > 0 Then
            alreadyUsed = False

            ' Check if this column number is already assigned in colDict
            Dim key As Variant
            For Each key In colDict.Keys
                If colDict(key) = col.Column Then
                    alreadyUsed = True
                    Exit For
                End If
            Next key

            ' If not already used, assign it
            If Not alreadyUsed Then
                GetColumnIndex = col.Column
                Exit Function
            End If
        End If
    Next col
End Function

Function SetupSAPScripting() As Boolean
    
    ' Create the SAP GUI scripting engine object
    On Error Resume Next
    Set SapGuiAuto = GetObject("SAPGUI")
    On Error GoTo ErrorHandler
    
    If Not IsObject(SapGuiAuto) Or SapGuiAuto Is Nothing Then
        SetupSAPScripting = False
        Exit Function
    End If
    
    On Error Resume Next
    Set SAPApplication = SapGuiAuto.GetScriptingEngine
    On Error GoTo ErrorHandler
    
    If Not IsObject(SAPApplication) Or SAPApplication Is Nothing Then
        SetupSAPScripting = False
        Exit Function
    End If
    
    ' Get the first connection and session
    On Error GoTo ErrorHandler
    Set Connection = SAPApplication.Children(0)
    Set session = Connection.Children(0)
    On Error GoTo ErrorHandler
    
    SetupSAPScripting = True
    
    If False Then
ErrorHandler:
    SetupSAPScripting = False
    End If
    
End Function

Function EndSAPScripting()
    ' Clean up
    Set session = Nothing
    Set Connection = Nothing
    Set SAPApplication = Nothing
    Set SapGuiAuto = Nothing
End Function

Function OptimizeCodeExecution(enable As Boolean)
    With Application
        If enable Then
            ' Disable settings for optimization
            .ScreenUpdating = False
            .Calculation = xlCalculationManual
            .EnableEvents = False
        Else
            ' Re-enable settings after optimization
            .ScreenUpdating = True
            .Calculation = xlCalculationAutomatic
            .EnableEvents = True
        End If
    End With
End Function

