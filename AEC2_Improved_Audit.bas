Attribute VB_Name = "AEC2_Improved_Audit"
Option Explicit

' Excel audit for AutoCAD Electrical-derived master sheets.
' Required sheets:
'   MasterComponentDetails
'   Master_ConnectionDetails
'
' This is a report-based audit. It does not call AutoCAD Electrical AEAUDIT.

Public Sub AEC2RunImprovedAudit()
    Dim wsC As Worksheet, wsN As Worksheet, wsR As Worksheet, wsCfg As Worksheet
    Dim oldCalc As XlCalculation

    oldCalc = Application.Calculation
    On Error GoTo Failed
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Application.Calculation = xlCalculationManual

    Set wsC = AEC2GetSheet("MasterComponentDetails")
    Set wsN = AEC2GetSheet("Master_ConnectionDetails")
    If wsC Is Nothing Then Err.Raise vbObjectError + 3101, , "MasterComponentDetails was not found."
    If wsN Is Nothing Then Err.Raise vbObjectError + 3102, , "Master_ConnectionDetails was not found."

    AEC2RequireHeaders wsC, Array("Tag", "Category", "Catalog Number", "Installation", "Location")
    AEC2RequireHeaders wsN, Array("Wire Number", "From Component", "From Pin", "To Component", "To Pin")

    Set wsCfg = AEC2GetConfig()
    Set wsR = AEC2GetResults()
    AEC2WriteHeaders wsR

    AEC2CheckWireNoConnection wsN, wsR
    AEC2CheckWireExceptions wsN, wsR
    AEC2CheckCableExceptions wsC, wsN, wsR
    AEC2CheckNoCatalog wsC, wsR
    AEC2CheckComponentDuplicates wsC, wsR, UCase$(Trim$(CStr(wsCfg.Range("B2").Value2)))
    AEC2CheckNoComponentConnection wsC, wsN, wsR
    AEC2CheckMixedNetworkAvailability wsC, wsR
    AEC2CheckTerminalDuplicates wsC, wsR
    AEC2CheckPinProxy wsN, wsR
    AEC2CheckContactsAvailability wsC, wsR

    AEC2FormatResults wsR
    MsgBox "Improved report-based audit completed. Review AEAudit_Results.", vbInformation

CleanExit:
    Application.Calculation = oldCalc
    Application.EnableEvents = True
    Application.ScreenUpdating = True
    Exit Sub

Failed:
    MsgBox "Audit stopped: " & Err.Description, vbCritical
    Resume CleanExit
End Sub

Private Function AEC2GetSheet(ByVal sheetName As String) As Worksheet
    On Error Resume Next
    Set AEC2GetSheet = ThisWorkbook.Worksheets(sheetName)
    On Error GoTo 0
End Function

Private Function AEC2GetConfig() As Worksheet
    Dim ws As Worksheet
    Set ws = AEC2GetSheet("AEAudit_Config")
    If ws Is Nothing Then
        Set ws = ThisWorkbook.Worksheets.Add(Before:=ThisWorkbook.Worksheets(1))
        ws.Name = "AEAudit_Config"
        ws.Range("A1:B4").Value = Array( _
            Array("Setting", "Value"), _
            Array("Component duplicate mode", "STANDARD"), _
            Array("Allowed duplicate modes", "STANDARD or INST_LOC"), _
            Array("Audit type", "REPORT-BASED PROXY"))
        ws.Columns("A:B").AutoFit
    End If
    Set AEC2GetConfig = ws
End Function

Private Function AEC2GetResults() As Worksheet
    Dim ws As Worksheet
    Set ws = AEC2GetSheet("AEAudit_Results")
    If ws Is Nothing Then
        Set ws = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))
        ws.Name = "AEAudit_Results"
    Else
        ws.Cells.Clear
    End If
    Set AEC2GetResults = ws
End Function

Private Sub AEC2WriteHeaders(ByVal ws As Worksheet)
    Dim headers As Variant, i As Long
    headers = Array("Check", "Severity", "Result", "Wire Number", "Cable Tag", _
        "Wire Color/Gauge", "Component Tag", "Category/Class", "Terminal Strip", "Terminal Number", _
        "Pin", "Installation", "Location", "Reference", "Drawing", "Error Message", _
        "Basis / Limitation", "Source Row(s)")
    For i = 0 To UBound(headers)
        ws.Cells(1, i + 1).Value = headers(i)
    Next i
End Sub

Private Sub AEC2AddResult(ByVal ws As Worksheet, ByVal checkName As String, ByVal severity As String, _
    ByVal resultType As String, Optional ByVal wireNumber As String = "", Optional ByVal cableTag As String = "", _
    Optional ByVal wireStyle As String = "", Optional ByVal componentTag As String = "", _
    Optional ByVal category As String = "", Optional ByVal terminalStrip As String = "", _
    Optional ByVal terminalNumber As String = "", Optional ByVal pin As String = "", _
    Optional ByVal installation As String = "", Optional ByVal location As String = "", _
    Optional ByVal reference As String = "", Optional ByVal drawing As String = "", _
    Optional ByVal message As String = "", Optional ByVal basis As String = "", _
    Optional ByVal sourceRows As String = "")

    Dim outputRow As Long
    outputRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row + 1
    ws.Cells(outputRow, 1).Resize(1, 18).Value = Array(checkName, severity, resultType, wireNumber, cableTag, _
        wireStyle, componentTag, category, terminalStrip, terminalNumber, pin, installation, location, _
        reference, drawing, message, basis, sourceRows)
End Sub

Private Function AEC2Column(ByVal ws As Worksheet, ByVal headerName As String) As Long
    Dim col As Long, lastCol As Long
    lastCol = ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column
    For col = 1 To lastCol
        If StrComp(Trim$(CStr(ws.Cells(1, col).Value2)), Trim$(headerName), vbTextCompare) = 0 Then
            AEC2Column = col
            Exit Function
        End If
    Next col
End Function

Private Sub AEC2RequireHeaders(ByVal ws As Worksheet, ByVal headers As Variant)
    Dim item As Variant
    For Each item In headers
        If AEC2Column(ws, CStr(item)) = 0 Then
            Err.Raise vbObjectError + 3110, , "Required header '" & CStr(item) & "' is missing from " & ws.Name & "."
        End If
    Next item
End Sub

Private Function AEC2Value(ByVal ws As Worksheet, ByVal rowNumber As Long, ByVal headerName As String) As String
    Dim col As Long, v As Variant
    col = AEC2Column(ws, headerName)
    If col = 0 Then Exit Function
    v = ws.Cells(rowNumber, col).Value2
    If Not IsError(v) And Not IsEmpty(v) Then AEC2Value = Trim$(CStr(v))
End Function

Private Function AEC2LastRow(ByVal ws As Worksheet) As Long
    Dim found As Range
    Set found = ws.Cells.Find("*", ws.Cells(1, 1), xlFormulas, xlPart, xlByRows, xlPrevious)
    If found Is Nothing Then AEC2LastRow = 1 Else AEC2LastRow = found.Row
End Function

Private Function AEC2IsUnknown(ByVal textValue As String) As Boolean
    textValue = UCase$(Trim$(textValue))
    AEC2IsUnknown = (textValue = "" Or textValue = "?" Or textValue = "??" Or textValue = "(??)")
End Function

Private Function AEC2NormalizePin(ByVal pinValue As String) As String
    pinValue = UCase$(Trim$(pinValue))
    If Right$(pinValue, 3) = ": I" Or Right$(pinValue, 3) = ": E" Then
        pinValue = Trim$(Left$(pinValue, Len(pinValue) - 3))
    End If
    AEC2NormalizePin = pinValue
End Function

Private Function AEC2Endpoint(ByVal inst As String, ByVal loc As String, ByVal tag As String, ByVal pin As String) As String
    AEC2Endpoint = UCase$(Trim$(inst) & "|" & Trim$(loc) & "|" & Trim$(tag) & "|" & AEC2NormalizePin(pin))
End Function

Private Function AEC2ComponentKey(ByVal inst As String, ByVal loc As String, ByVal tag As String) As String
    AEC2ComponentKey = UCase$(Trim$(inst) & "|" & Trim$(loc) & "|" & Trim$(tag))
End Function

Private Function AEC2IsTerminal(ByVal ws As Worksheet, ByVal rowNumber As Long) As Boolean
    AEC2IsTerminal = (UCase$(AEC2Value(ws, rowNumber, "Category")) = "TRMS" Or _
        (AEC2Value(ws, rowNumber, "Terminal Strip") <> "" And AEC2Value(ws, rowNumber, "Terminal Number") <> ""))
End Function

Private Function AEC2IsCableComponent(ByVal ws As Worksheet, ByVal rowNumber As Long) As Boolean
    AEC2IsCableComponent = (InStr(1, UCase$(AEC2Value(ws, rowNumber, "Component Status")), "CABLE", vbTextCompare) > 0)
End Function

Private Function AEC2IsConsolidatedGroup(ByVal ws As Worksheet, ByVal rowNumber As Long) As Boolean
    Dim sourceText As String, tagText As String
    sourceText = UCase$(AEC2Value(ws, rowNumber, "Source Reports"))
    tagText = UCase$(AEC2Value(ws, rowNumber, "Tag"))
    AEC2IsConsolidatedGroup = ((tagText = "CR" Or tagText = "FU") And InStr(1, sourceText, "PANEL", vbTextCompare) > 0)
End Function

' 1. Wire - No Connection
Private Sub AEC2CheckWireNoConnection(ByVal ws As Worksheet, ByVal outputWs As Worksheet)
    Dim rowNumber As Long, fromTag As String, toTag As String
    For rowNumber = 2 To AEC2LastRow(ws)
        fromTag = AEC2Value(ws, rowNumber, "From Component")
        toTag = AEC2Value(ws, rowNumber, "To Component")
        If AEC2IsUnknown(fromTag) Xor AEC2IsUnknown(toTag) Then
            AEC2AddResult outputWs, "Wire - No Connection", "WARNING", "POSSIBLE", _
                wireNumber:=AEC2Value(ws, rowNumber, "Wire Number"), _
                cableTag:=AEC2Value(ws, rowNumber, "Cable"), _
                wireStyle:=AEC2Value(ws, rowNumber, "Wire Type"), _
                message:="Only one component endpoint is present in the report row.", _
                basis:="Report-based proxy only. A blank endpoint can represent an open end, shield/drain, bus, or incomplete report data.", _
                sourceRows:=CStr(rowNumber)
        End If
    Next rowNumber
End Sub

' 2. Wire Exception
' Missing wire numbers are direct checks.
' Duplicate wire numbers are detected by building connected endpoint groups first.
Private Sub AEC2CheckWireExceptions(ByVal ws As Worksheet, ByVal outputWs As Worksheet)
    Dim parent As Object, rank As Object, rowWire As Object, rowA As Object, rowB As Object
    Dim rowNumber As Long, endpointA As String, endpointB As String, wireNumber As String
    Dim networksByWire As Object, rowsByWireNetwork As Object, networkKey As String, rootKey As String
    Dim key As Variant, parts As Variant, summary As String, networkSet As Object

    Set parent = CreateObject("Scripting.Dictionary"): parent.CompareMode = vbTextCompare
    Set rank = CreateObject("Scripting.Dictionary"): rank.CompareMode = vbTextCompare
    Set rowWire = CreateObject("Scripting.Dictionary")
    Set rowA = CreateObject("Scripting.Dictionary")
    Set rowB = CreateObject("Scripting.Dictionary")

    For rowNumber = 2 To AEC2LastRow(ws)
        endpointA = AEC2Endpoint(AEC2Value(ws, rowNumber, "From Installation"), AEC2Value(ws, rowNumber, "From Location"), _
            AEC2Value(ws, rowNumber, "From Component"), AEC2Value(ws, rowNumber, "From Pin"))
        endpointB = AEC2Endpoint(AEC2Value(ws, rowNumber, "To Installation"), AEC2Value(ws, rowNumber, "To Location"), _
            AEC2Value(ws, rowNumber, "To Component"), AEC2Value(ws, rowNumber, "To Pin"))
        wireNumber = AEC2Value(ws, rowNumber, "Wire Number")

        If wireNumber = "" Then
            If Not (AEC2IsUnknown(AEC2Value(ws, rowNumber, "From Component")) And _
                    AEC2IsUnknown(AEC2Value(ws, rowNumber, "To Component"))) Then
                AEC2AddResult outputWs, "Wire Exception", "ERROR", "FLAG", _
                    message:="Connection row has no wire number.", _
                    basis:="Direct blank Wire Number test.", sourceRows:=CStr(rowNumber)
            End If
        End If

        rowWire(CStr(rowNumber)) = wireNumber
        rowA(CStr(rowNumber)) = endpointA
        rowB(CStr(rowNumber)) = endpointB

        If Not AEC2EndpointIsEmpty(endpointA) Then AEC2UFAdd parent, rank, endpointA
        If Not AEC2EndpointIsEmpty(endpointB) Then AEC2UFAdd parent, rank, endpointB
        If Not AEC2EndpointIsEmpty(endpointA) And Not AEC2EndpointIsEmpty(endpointB) Then
            AEC2UFUnion parent, rank, endpointA, endpointB
        End If
    Next rowNumber

    Set networksByWire = CreateObject("Scripting.Dictionary"): networksByWire.CompareMode = vbTextCompare
    Set rowsByWireNetwork = CreateObject("Scripting.Dictionary"): rowsByWireNetwork.CompareMode = vbTextCompare

    For rowNumber = 2 To AEC2LastRow(ws)
        wireNumber = CStr(rowWire(CStr(rowNumber)))
        If wireNumber <> "" Then
            endpointA = CStr(rowA(CStr(rowNumber)))
            endpointB = CStr(rowB(CStr(rowNumber)))
            If Not AEC2EndpointIsEmpty(endpointA) Then
                rootKey = AEC2UFFind(parent, endpointA)
            ElseIf Not AEC2EndpointIsEmpty(endpointB) Then
                rootKey = AEC2UFFind(parent, endpointB)
            Else
                rootKey = "ROW:" & CStr(rowNumber)
            End If

            networkKey = UCase$(wireNumber) & Chr$(30) & rootKey
            If Not networksByWire.Exists(UCase$(wireNumber)) Then
                Set networkSet = CreateObject("Scripting.Dictionary")
                networkSet.CompareMode = vbTextCompare
                networksByWire.Add UCase$(wireNumber), networkSet
            End If
            networksByWire(UCase$(wireNumber))(rootKey) = True

            If rowsByWireNetwork.Exists(networkKey) Then
                rowsByWireNetwork(networkKey) = rowsByWireNetwork(networkKey) & ", " & CStr(rowNumber)
            Else
                rowsByWireNetwork.Add networkKey, CStr(rowNumber)
            End If
        End If
    Next rowNumber

    For Each key In networksByWire.Keys
        If networksByWire(key).Count > 1 Then
            summary = ""
            For Each parts In networksByWire(key).Keys
                networkKey = CStr(key) & Chr$(30) & CStr(parts)
                If summary <> "" Then summary = summary & " | "
                summary = summary & "Network rows " & CStr(rowsByWireNetwork(networkKey))
            Next parts
            AEC2AddResult outputWs, "Wire Exception", "WARNING", "POSSIBLE", _
                wireNumber:=CStr(key), _
                message:="The same wire number appears on more than one disconnected endpoint group.", _
                basis:="Connected groups were reconstructed from Installation + Location + Component + Pin. Signal-arrow links and drawing topology are not available.", _
                sourceRows:=summary
        End If
    Next key
End Sub

Private Function AEC2EndpointIsEmpty(ByVal endpointKey As String) As Boolean
    Dim parts As Variant
    parts = Split(endpointKey, "|")
    AEC2EndpointIsEmpty = (UBound(parts) < 3 Or (parts(0) = "" And parts(1) = "" And AEC2IsUnknown(parts(2)) And parts(3) = ""))
End Function

Private Sub AEC2UFAdd(ByVal parent As Object, ByVal rank As Object, ByVal node As String)
    If Not parent.Exists(node) Then
        parent.Add node, node
        rank.Add node, 0
    End If
End Sub

Private Function AEC2UFFind(ByVal parent As Object, ByVal node As String) As String
    Dim p As String
    p = CStr(parent(node))
    If StrComp(p, node, vbTextCompare) <> 0 Then
        parent(node) = AEC2UFFind(parent, p)
    End If
    AEC2UFFind = CStr(parent(node))
End Function

Private Sub AEC2UFUnion(ByVal parent As Object, ByVal rank As Object, ByVal nodeA As String, ByVal nodeB As String)
    Dim rootA As String, rootB As String, rankA As Long, rankB As Long
    rootA = AEC2UFFind(parent, nodeA)
    rootB = AEC2UFFind(parent, nodeB)
    If StrComp(rootA, rootB, vbTextCompare) = 0 Then Exit Sub
    rankA = CLng(rank(rootA)): rankB = CLng(rank(rootB))
    If rankA < rankB Then
        parent(rootA) = rootB
    ElseIf rankA > rankB Then
        parent(rootB) = rootA
    Else
        parent(rootB) = rootA
        rank(rootA) = rankA + 1
    End If
End Sub

' 3. Cable Exception
Private Sub AEC2CheckCableExceptions(ByVal comp As Worksheet, ByVal conn As Worksheet, ByVal outputWs As Worksheet)
    Dim tags As Object, conductors As Object
    Dim rowNumber As Long, cableTag As String, colorValue As String, gaugeValue As String, styleValue As String
    Dim key As String

    Set tags = CreateObject("Scripting.Dictionary"): tags.CompareMode = vbTextCompare
    Set conductors = CreateObject("Scripting.Dictionary"): conductors.CompareMode = vbTextCompare

    For rowNumber = 2 To AEC2LastRow(comp)
        If AEC2IsCableComponent(comp, rowNumber) Then
            cableTag = AEC2Value(comp, rowNumber, "Tag")
            If cableTag <> "" Then
                If tags.Exists(cableTag) Then
                    AEC2AddResult outputWs, "Cable Exception", "ERROR", "FLAG", cableTag:=cableTag, _
                        message:="Cable tag appears on more than one cable master row.", _
                        basis:="Duplicate cable-tag rows in the current master sheet.", _
                        sourceRows:=CStr(tags(cableTag)) & ", " & CStr(rowNumber)
                Else
                    tags.Add cableTag, rowNumber
                End If
            End If
        End If
    Next rowNumber

    For rowNumber = 2 To AEC2LastRow(conn)
        cableTag = AEC2Value(conn, rowNumber, "Cable")
        colorValue = AEC2Value(conn, rowNumber, "Wire Color")
        gaugeValue = AEC2Value(conn, rowNumber, "Wire Gauge")
        styleValue = AEC2Value(conn, rowNumber, "Wire Type")

        If cableTag <> "" Then
            If colorValue <> "" Or gaugeValue <> "" Then
                key = UCase$(cableTag & "|" & colorValue & "|" & gaugeValue)
                styleValue = Trim$(colorValue & IIf(colorValue <> "" And gaugeValue <> "", " / ", "") & gaugeValue)
            ElseIf styleValue <> "" Then
                key = UCase$(cableTag & "|TYPE|" & styleValue)
            Else
                key = ""
            End If

            If key <> "" Then
                If conductors.Exists(key) Then
                    AEC2AddResult outputWs, "Cable Exception", "WARNING", "POSSIBLE", _
                        cableTag:=cableTag, wireStyle:=styleValue, _
                        message:="The same conductor color/gauge identifier repeats within the cable.", _
                        basis:="Uses Wire Color + Wire Gauge when available; otherwise falls back to Wire Type. Repeated shield/drain identifiers can be intentional.", _
                        sourceRows:=CStr(conductors(key)) & ", " & CStr(rowNumber)
                Else
                    conductors.Add key, rowNumber
                End If
            End If
        End If
    Next rowNumber
End Sub

' 4. Component - No Catalog Number
Private Sub AEC2CheckNoCatalog(ByVal ws As Worksheet, ByVal outputWs As Worksheet)
    Dim rowNumber As Long, tagValue As String
    For rowNumber = 2 To AEC2LastRow(ws)
        tagValue = AEC2Value(ws, rowNumber, "Tag")
        If tagValue <> "" And Not AEC2IsTerminal(ws, rowNumber) And Not AEC2IsConsolidatedGroup(ws, rowNumber) Then
            If AEC2Value(ws, rowNumber, "Catalog Number") = "" Then
                AEC2AddResult outputWs, "Component - No Catalog Number", "ERROR", "FLAG", _
                    componentTag:=tagValue, category:=AEC2Value(ws, rowNumber, "Category"), _
                    installation:=AEC2Value(ws, rowNumber, "Installation"), location:=AEC2Value(ws, rowNumber, "Location"), _
                    message:="Component has no primary catalog number.", basis:="Direct blank Catalog Number test.", _
                    sourceRows:=CStr(rowNumber)
            End If
        End If
    Next rowNumber
End Sub

' 5. Component Duplication
Private Sub AEC2CheckComponentDuplicates(ByVal ws As Worksheet, ByVal outputWs As Worksheet, ByVal mode As String)
    Dim seen As Object, rowNumber As Long, tagValue As String, key As String
    Set seen = CreateObject("Scripting.Dictionary"): seen.CompareMode = vbTextCompare

    If mode <> "STANDARD" And mode <> "INST_LOC" Then mode = "STANDARD"

    For rowNumber = 2 To AEC2LastRow(ws)
        tagValue = AEC2Value(ws, rowNumber, "Tag")
        If tagValue <> "" And Not AEC2IsTerminal(ws, rowNumber) And Not AEC2IsCableComponent(ws, rowNumber) And Not AEC2IsConsolidatedGroup(ws, rowNumber) Then
            If mode = "INST_LOC" Then
                key = UCase$(tagValue & "|" & AEC2Value(ws, rowNumber, "Installation") & "|" & AEC2Value(ws, rowNumber, "Location"))
            Else
                key = UCase$(tagValue)
            End If

            If seen.Exists(key) Then
                AEC2AddResult outputWs, "Component Duplication", "ERROR", "FLAG", _
                    componentTag:=tagValue, category:=AEC2Value(ws, rowNumber, "Category"), _
                    installation:=AEC2Value(ws, rowNumber, "Installation"), location:=AEC2Value(ws, rowNumber, "Location"), _
                    message:="Duplicate component identity.", _
                    basis:="Compared using " & mode & " mode. Consolidated source rows can still hide original duplicates.", _
                    sourceRows:=CStr(seen(key)) & ", " & CStr(rowNumber)
            Else
                seen.Add key, rowNumber
            End If
        End If
    Next rowNumber
End Sub

' 6. Component - No Connection
' Uses full component identity rather than tag alone and excludes cables/consolidated groups.
Private Sub AEC2CheckNoComponentConnection(ByVal comp As Worksheet, ByVal conn As Worksheet, ByVal outputWs As Worksheet)
    Dim connected As Object, rowNumber As Long, key As String, tagValue As String
    Set connected = CreateObject("Scripting.Dictionary"): connected.CompareMode = vbTextCompare

    For rowNumber = 2 To AEC2LastRow(conn)
        tagValue = AEC2Value(conn, rowNumber, "From Component")
        If Not AEC2IsUnknown(tagValue) Then
            key = AEC2ComponentKey(AEC2Value(conn, rowNumber, "From Installation"), AEC2Value(conn, rowNumber, "From Location"), tagValue)
            connected(key) = True
        End If
        tagValue = AEC2Value(conn, rowNumber, "To Component")
        If Not AEC2IsUnknown(tagValue) Then
            key = AEC2ComponentKey(AEC2Value(conn, rowNumber, "To Installation"), AEC2Value(conn, rowNumber, "To Location"), tagValue)
            connected(key) = True
        End If
    Next rowNumber

    For rowNumber = 2 To AEC2LastRow(comp)
        tagValue = AEC2Value(comp, rowNumber, "Tag")
        If tagValue <> "" And Not AEC2IsTerminal(comp, rowNumber) And Not AEC2IsCableComponent(comp, rowNumber) And Not AEC2IsConsolidatedGroup(comp, rowNumber) Then
            key = AEC2ComponentKey(AEC2Value(comp, rowNumber, "Installation"), AEC2Value(comp, rowNumber, "Location"), tagValue)
            If Not connected.Exists(key) Then
                AEC2AddResult outputWs, "Component - No Connection", "WARNING", "POSSIBLE", _
                    componentTag:=tagValue, category:=AEC2Value(comp, rowNumber, "Category"), _
                    installation:=AEC2Value(comp, rowNumber, "Installation"), location:=AEC2Value(comp, rowNumber, "Location"), _
                    message:="Component identity is absent from all From/To endpoints.", _
                    basis:="Report-based isolation proxy. It cannot verify each X?TERMn connection attribute on the original block.", _
                    sourceRows:=CStr(rowNumber)
            End If
        End If
    Next rowNumber
End Sub

' 7. Mixed Component Network
Private Sub AEC2CheckMixedNetworkAvailability(ByVal ws As Worksheet, ByVal outputWs As Worksheet)
    If AEC2Column(ws, "WDTYPE") = 0 Then
        AEC2AddResult outputWs, "Mixed Component Network", "INFO", "LIMITED", _
            message:="Mixed network check was not run.", _
            basis:="MasterComponentDetails does not contain WDTYPE. Add WDTYPE and join components to the reconstructed wire networks."
    Else
        AEC2AddResult outputWs, "Mixed Component Network", "INFO", "LIMITED", _
            message:="WDTYPE exists, but the current module does not assume that the consolidated component rows preserve every original symbol instance.", _
            basis:="Run this check against unconsolidated Component Report rows joined to the reconstructed networks."
    End If
End Sub

' 8. Terminal Duplication
Private Sub AEC2CheckTerminalDuplicates(ByVal ws As Worksheet, ByVal outputWs As Worksheet)
    Dim seen As Object, rowNumber As Long, stripValue As String, terminalValue As String, key As String
    Set seen = CreateObject("Scripting.Dictionary"): seen.CompareMode = vbTextCompare

    For rowNumber = 2 To AEC2LastRow(ws)
        If AEC2IsTerminal(ws, rowNumber) Then
            stripValue = AEC2Value(ws, rowNumber, "Terminal Strip")
            terminalValue = AEC2NormalizePin(AEC2Value(ws, rowNumber, "Terminal Number"))
            If stripValue <> "" And terminalValue <> "" Then
                key = UCase$(AEC2Value(ws, rowNumber, "Installation") & "|" & AEC2Value(ws, rowNumber, "Location") & "|" & stripValue & "|" & terminalValue)
                If seen.Exists(key) Then
                    AEC2AddResult outputWs, "Terminal Duplication", "ERROR", "FLAG", _
                        terminalStrip:=stripValue, terminalNumber:=terminalValue, _
                        installation:=AEC2Value(ws, rowNumber, "Installation"), location:=AEC2Value(ws, rowNumber, "Location"), _
                        message:="Duplicate terminal number on the same terminal strip identity.", _
                        basis:="Compared as Installation + Location + Terminal Strip + Terminal Number. WDN exclusions are not available.", _
                        sourceRows:=CStr(seen(key)) & ", " & CStr(rowNumber)
                Else
                    seen.Add key, rowNumber
                End If
            End If
        End If
    Next rowNumber
End Sub

' 9. Pin Exception
' The connection report cannot prove duplicate TERMn attributes on a block.
' This routine remains a conservative peer-endpoint proxy and ignores exact duplicate rows.
Private Sub AEC2CheckPinProxy(ByVal ws As Worksheet, ByVal outputWs As Worksheet)
    Dim peers As Object, rowNumber As Long
    Set peers = CreateObject("Scripting.Dictionary"): peers.CompareMode = vbTextCompare
    For rowNumber = 2 To AEC2LastRow(ws)
        AEC2ReviewPinSide ws, outputWs, peers, rowNumber, True
        AEC2ReviewPinSide ws, outputWs, peers, rowNumber, False
    Next rowNumber
End Sub

Private Sub AEC2ReviewPinSide(ByVal ws As Worksheet, ByVal outputWs As Worksheet, ByVal peers As Object, _
    ByVal rowNumber As Long, ByVal fromSide As Boolean)

    Dim inst As String, loc As String, tagValue As String, pinValue As String, peer As String, key As String
    Dim peerSet As Object

    If fromSide Then
        inst = AEC2Value(ws, rowNumber, "From Installation")
        loc = AEC2Value(ws, rowNumber, "From Location")
        tagValue = AEC2Value(ws, rowNumber, "From Component")
        pinValue = AEC2NormalizePin(AEC2Value(ws, rowNumber, "From Pin"))
        peer = AEC2Endpoint(AEC2Value(ws, rowNumber, "To Installation"), AEC2Value(ws, rowNumber, "To Location"), _
            AEC2Value(ws, rowNumber, "To Component"), AEC2Value(ws, rowNumber, "To Pin"))
    Else
        inst = AEC2Value(ws, rowNumber, "To Installation")
        loc = AEC2Value(ws, rowNumber, "To Location")
        tagValue = AEC2Value(ws, rowNumber, "To Component")
        pinValue = AEC2NormalizePin(AEC2Value(ws, rowNumber, "To Pin"))
        peer = AEC2Endpoint(AEC2Value(ws, rowNumber, "From Installation"), AEC2Value(ws, rowNumber, "From Location"), _
            AEC2Value(ws, rowNumber, "From Component"), AEC2Value(ws, rowNumber, "From Pin"))
    End If

    If AEC2IsUnknown(tagValue) Or pinValue = "" Then Exit Sub
    key = AEC2Endpoint(inst, loc, tagValue, pinValue)

    If Not peers.Exists(key) Then
        Set peerSet = CreateObject("Scripting.Dictionary"): peerSet.CompareMode = vbTextCompare
        peers.Add key, peerSet
    End If

    If Not peers(key).Exists(peer) Then
        peers(key).Add peer, rowNumber
        If peers(key).Count > 1 Then
            AEC2AddResult outputWs, "Pin Exception", "WARNING", "POSSIBLE", _
                wireNumber:=AEC2Value(ws, rowNumber, "Wire Number"), componentTag:=tagValue, pin:=pinValue, _
                installation:=inst, location:=loc, _
                message:="The same component/pin appears with more than one peer endpoint in the connection report.", _
                basis:="This can be a legitimate network branch. Native pin exception requires distinct TERMn/X?TERMn attributes from the source symbol.", _
                sourceRows:=AEC2DictionaryValues(peers(key))
        End If
    End If
End Sub

Private Function AEC2DictionaryValues(ByVal dictionaryObject As Object) As String
    Dim key As Variant, textValue As String
    For Each key In dictionaryObject.Keys
        If textValue <> "" Then textValue = textValue & ", "
        textValue = textValue & CStr(dictionaryObject(key))
    Next key
    AEC2DictionaryValues = textValue
End Function

' 10. Contacts
Private Sub AEC2CheckContactsAvailability(ByVal ws As Worksheet, ByVal outputWs As Worksheet)
    If AEC2Column(ws, "TAG2") = 0 Then
        AEC2AddResult outputWs, "Contacts", "INFO", "LIMITED", _
            message:="Orphan-child contact check was not run.", _
            basis:="The current master component sheet does not contain TAG2 or another verified parent-child relationship field."
    Else
        AEC2AddResult outputWs, "Contacts", "INFO", "LIMITED", _
            message:="TAG2 exists, but parent/child role and original symbol-instance data must be validated before flagging contacts.", _
            basis:="Use unconsolidated child rows and confirm that each child parent reference matches a parent TAG1 identity."
    End If
End Sub

Private Sub AEC2FormatResults(ByVal ws As Worksheet)
    Dim rowNumber As Long, lastRow As Long, resultValue As String
    With ws.Rows(1)
        .Font.Bold = True
        .Font.Color = vbWhite
        .Interior.Color = RGB(31, 78, 121)
    End With

    ws.Range("A1:R1").AutoFilter
    ws.Columns("A:R").AutoFit
    ws.Columns("P:Q").ColumnWidth = 58
    ws.Columns("P:R").WrapText = True
    lastRow = AEC2LastRow(ws)

    For rowNumber = 2 To lastRow
        resultValue = UCase$(AEC2Value(ws, rowNumber, "Result"))
        Select Case resultValue
            Case "FLAG": ws.Rows(rowNumber).Interior.Color = RGB(255, 199, 206)
            Case "POSSIBLE": ws.Rows(rowNumber).Interior.Color = RGB(255, 235, 156)
            Case "LIMITED": ws.Rows(rowNumber).Interior.Color = RGB(221, 235, 247)
        End Select
    Next rowNumber

    ws.Activate
    ActiveWindow.FreezePanes = False
    ActiveWindow.SplitRow = 1
    ActiveWindow.FreezePanes = True
End Sub
