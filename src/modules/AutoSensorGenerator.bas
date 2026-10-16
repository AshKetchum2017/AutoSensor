Option Explicit

Private Const AS_AREA_WIDTH As Double = 325#
Private Const AS_AREA_HEIGHT As Double = 485#
Private Const AS_POSITION_TOLERANCE As Double = 0.000001

Public Sub ASGenerate(ByVal whiteSpace As Boolean, ByVal sensorText As String)
    Dim doc As Document
    Dim targetPage As Page
    Dim previousLayer As Layer
    Dim previousSelection As ShapeRange
    Dim layers(1 To 3) As Layer
    Dim wasEditable(1 To 3) As Boolean
    Dim previousUnit As cdrUnit
    Dim previousHUnit As cdrUnit
    Dim previousVUnit As cdrUnit
    Dim commandOpen As Boolean
    Dim stateSaved As Boolean
    Dim targetCreated As Boolean
    Dim created As ShapeRange
    Dim result As Shape
    Dim newShape As Shape
    Dim shapeTracked As Boolean
    Dim anchor As Shape
    Dim positionsX(1 To 5) As Double
    Dim positionsY(1 To 5) As Double
    Dim originX As Double, originY As Double
    Dim diameter As Double
    Dim itemCount As Long, i As Long
    Dim operation As String
    Dim errorNumber As Long, errorSource As String, errorDescription As String
    Dim cleanupIssues As String

    On Error GoTo GenerateFailed
    operation = "Memvalidasi dokumen dan ukuran"
    If Application.Documents.Count = 0 Then
        Err.Raise vbObjectError + 6100, "AutoSensor", "Tidak ada dokumen aktif."
    End If
    If Not whiteSpace Then diameter = ASReadDiameter(sensorText)
    Set doc = ActiveDocument
    Set targetPage = doc.ActivePage
    Set previousLayer = doc.ActiveLayer
    Set previousSelection = doc.SelectionRange
    Set created = Application.CreateShapeRange

    operation = "Mencari layer pada halaman aktif"
    Set layers(1) = ASFindLayer(targetPage, "Layer 1", False)
    If whiteSpace Then
        Set layers(2) = ASFindLayer(targetPage, "scpro2_regmarks", True)
        Set layers(3) = ASFindLayer(targetPage, "scpro2_printonly", True)
    End If
    previousUnit = doc.Unit
    previousHUnit = doc.Rulers.HUnits
    previousVUnit = doc.Rulers.VUnits
    For i = 1 To 3
        If Not layers(i) Is Nothing Then wasEditable(i) = layers(i).Editable
    Next i
    stateSaved = True

    operation = "Memulai CommandGroup"
    doc.BeginCommandGroup "AutoSensor"
    commandOpen = True
    doc.Unit = cdrMillimeter

    ' Page.CenterX/CenterY mengikuti origin dokumen; area virtual tetap di tengah page.
    originX = targetPage.CenterX - AS_AREA_WIDTH / 2#
    originY = targetPage.CenterY - AS_AREA_HEIGHT / 2#
    operation = "Membuka sementara layer anchor"
    For i = 2 To 3
        If Not layers(i) Is Nothing Then
            If Not layers(i).Editable Then layers(i).Editable = True
        End If
    Next i

    If whiteSpace Then
        positionsX(1) = 6#: positionsY(1) = 10#
        positionsX(2) = 6#: positionsY(2) = 475#
        positionsX(3) = 319#: positionsY(3) = 10#
        positionsX(4) = 319#: positionsY(4) = 475#
        positionsX(5) = 162.5: positionsY(5) = 11.5
        itemCount = 5
        ' Seluruh anchor wajib valid sebelum objek pertama dibuat.
        For i = 1 To itemCount
            operation = "Mencari anchor " & CStr(i)
            If i = 5 Then
                Set anchor = ASFindAnchor(layers(3), originX + positionsX(i), _
                                          originY + positionsY(i), 12#, "qr.svg")
            Else
                Set anchor = ASFindAnchor(layers(2), originX + positionsX(i), _
                                          originY + positionsY(i), 10#, vbNullString)
            End If
            positionsX(i) = anchor.CenterX
            positionsY(i) = anchor.CenterY
        Next i
        diameter = 5.5
    Else
        itemCount = 4
        If diameter <= 5# Then
            positionsX(1) = 10#: positionsY(1) = 10#
            positionsX(2) = 10#: positionsY(2) = 475#
            positionsX(3) = 315#: positionsY(3) = 10#
            positionsX(4) = 315#: positionsY(4) = 475#
        Else
            positionsX(1) = 11#: positionsY(1) = 11#
            positionsX(2) = 11#: positionsY(2) = 474#
            positionsX(3) = 314#: positionsY(3) = 11#
            positionsX(4) = 314#: positionsY(4) = 474#
        End If
        For i = 1 To itemCount
            positionsX(i) = originX + positionsX(i)
            positionsY(i) = originY + positionsY(i)
        Next i
    End If

    operation = "Menyiapkan Layer 1"
    If layers(1) Is Nothing Then
        Set layers(1) = targetPage.CreateLayer("Layer 1")
        targetCreated = True
        wasEditable(1) = layers(1).Editable
    End If
    If Not layers(1).Editable Then layers(1).Editable = True

    For i = 1 To itemCount
        operation = "Membuat objek " & CStr(i)
        Set newShape = Nothing
        shapeTracked = False
        If whiteSpace And i = 5 Then
            Set newShape = layers(1).CreateRectangle(positionsX(i) - 5.25, _
                positionsY(i) + 5.25, positionsX(i) + 5.25, positionsY(i) - 5.25)
        Else
            Set newShape = layers(1).CreateEllipse2(positionsX(i), positionsY(i), _
                                                  diameter / 2#, diameter / 2#)
        End If
        ' Daftarkan sebelum styling agar error fill/outline tetap dapat dibersihkan.
        created.Add newShape
        shapeTracked = True
        newShape.Outline.SetNoOutline
        If whiteSpace Then
            newShape.Fill.ApplyUniformFill Application.CreateRGBColor(255, 255, 255)
        Else
            newShape.Fill.ApplyUniformFill Application.CreateRGBColor(0, 0, 0)
        End If
    Next i

    operation = "Mengelompokkan hasil di bagian teratas Layer 1"
    layers(1).Activate
    Set result = created.Group
    result.MoveToLayer layers(1)
    result.OrderToFront

    operation = "Menyeleksi group hasil"
    result.CreateSelection

    operation = "Memulihkan state dokumen"
    ASRestoreState doc, previousLayer, layers, wasEditable, _
                   previousUnit, previousHUnit, previousVUnit
    operation = "Menutup CommandGroup"
    doc.EndCommandGroup
    commandOpen = False
    Exit Sub

GenerateFailed:
    errorNumber = Err.Number
    errorSource = Err.Source
    errorDescription = Err.Description
    Resume GenerateCleanup

GenerateCleanup:
    On Error Resume Next
    ' CommandGroup mengelompokkan Undo; hapus hanya hasil baru saat gagal.
    If commandOpen Then
        If Not layers(1) Is Nothing Then
            Err.Clear
            layers(1).Editable = True
            ASRecordCleanupIssue cleanupIssues, "membuka layer hasil"
        End If
        If Not result Is Nothing Then
            Err.Clear
            result.Delete
            If Err.Number = 0 Then Set created = Nothing
            ASRecordCleanupIssue cleanupIssues, "menghapus group hasil"
        End If
        If Not shapeTracked And Not newShape Is Nothing Then
            Err.Clear
            newShape.Delete
            ASRecordCleanupIssue cleanupIssues, "menghapus objek yang belum terdaftar"
        End If
        If Not created Is Nothing Then
            Err.Clear
            If created.Count > 0 Then created.Delete
            ASRecordCleanupIssue cleanupIssues, "menghapus objek hasil"
        End If
        If targetCreated Then
            Err.Clear
            layers(1).Delete
            ASRecordCleanupIssue cleanupIssues, "menghapus layer hasil baru"
            Set layers(1) = Nothing
        End If
        If stateSaved Then
            ' Coba tiap pemulihan secara terpisah, meskipun pemulihan lain gagal.
            For i = 1 To 3
                If Not layers(i) Is Nothing Then
                    Err.Clear
                    layers(i).Editable = wasEditable(i)
                    ASRecordCleanupIssue cleanupIssues, "memulihkan kunci layer " & CStr(i)
                End If
            Next i
            Err.Clear
            previousLayer.Activate
            ASRecordCleanupIssue cleanupIssues, "memulihkan layer aktif"
            Err.Clear
            doc.ClearSelection
            ASRecordCleanupIssue cleanupIssues, "membersihkan selection hasil"
            Err.Clear
            If previousSelection.Count > 0 Then previousSelection.CreateSelection
            ASRecordCleanupIssue cleanupIssues, "memulihkan selection"
            Err.Clear
            doc.Unit = previousUnit
            ASRecordCleanupIssue cleanupIssues, "memulihkan unit dokumen"
            Err.Clear
            doc.Rulers.HUnits = previousHUnit
            ASRecordCleanupIssue cleanupIssues, "memulihkan unit ruler horizontal"
            Err.Clear
            doc.Rulers.VUnits = previousVUnit
            ASRecordCleanupIssue cleanupIssues, "memulihkan unit ruler vertikal"
        End If
        Err.Clear
        doc.EndCommandGroup
        ASRecordCleanupIssue cleanupIssues, "menutup CommandGroup"
    End If
    On Error GoTo 0
    MsgBox "AutoSensor dibatalkan." & vbCrLf & operation & ": " & errorDescription & _
           vbCrLf & "Error " & CStr(errorNumber) & " (" & errorSource & ")" & _
           cleanupIssues, vbExclamation, "AutoSensor"
End Sub

Private Function ASReadDiameter(ByVal text As String) As Double
    Dim normalized As String
    Dim i As Long, decimalCount As Long, digitCount As Long
    Dim character As String

    normalized = Replace(Trim$(text), ",", ".")
    For i = 1 To Len(normalized)
        character = Mid$(normalized, i, 1)
        If character = "." Then
            decimalCount = decimalCount + 1
        ElseIf character >= "0" And character <= "9" Then
            digitCount = digitCount + 1
        Else
            Err.Raise vbObjectError + 6101, "AutoSensor", _
                      "Ukuran sensor harus angka positif dalam mm, tanpa satuan/pemisah ribuan."
        End If
    Next i
    If digitCount = 0 Or decimalCount > 1 Then
        Err.Raise vbObjectError + 6101, "AutoSensor", "Masukkan ukuran sensor yang valid dalam mm."
    End If
    ASReadDiameter = Val(normalized)
    If ASReadDiameter <= 0# Then
        Err.Raise vbObjectError + 6101, "AutoSensor", "Ukuran sensor harus lebih dari 0 mm."
    End If
End Function

Private Function ASFindLayer(ByVal targetPage As Page, ByVal layerName As String, _
                             ByVal required As Boolean) As Layer
    Dim candidate As Layer
    For Each candidate In targetPage.Layers
        If StrComp(candidate.Name, layerName, vbTextCompare) = 0 Then
            If Not ASFindLayer Is Nothing Then
                Err.Raise vbObjectError + 6102, "AutoSensor", _
                          "Ditemukan lebih dari satu layer bernama " & layerName & "."
            End If
            Set ASFindLayer = candidate
        End If
    Next candidate
    If required And ASFindLayer Is Nothing Then
        Err.Raise vbObjectError + 6103, "AutoSensor", "Layer tidak ditemukan: " & layerName & "."
    End If
End Function

Private Function ASFindAnchor(ByVal sourceLayer As Layer, ByVal centerX As Double, _
                              ByVal centerY As Double, ByVal searchSize As Double, _
                              ByVal requiredName As String) As Shape
    Dim candidate As Shape
    Dim candidateCount As Long
    Dim halfSize As Double
    halfSize = searchSize / 2# + AS_POSITION_TOLERANCE
    ' Hanya objek langsung pada layer: group dihitung satu, tidak menelusuri anggotanya.
    For Each candidate In sourceLayer.Shapes
        If Len(requiredName) = 0 Or StrComp(candidate.Name, requiredName, vbTextCompare) = 0 Then
            If Abs(candidate.CenterX - centerX) <= halfSize And _
               Abs(candidate.CenterY - centerY) <= halfSize Then
                candidateCount = candidateCount + 1
                Set ASFindAnchor = candidate
            End If
        End If
    Next candidate
    If candidateCount <> 1 Then
        Err.Raise vbObjectError + 6104, "AutoSensor", _
                  "Anchor pada layer " & sourceLayer.Name & " (area " & CStr(searchSize) & _
                  " mm, nama: " & IIf(Len(requiredName) = 0, "bebas", requiredName) & _
                  ") harus tepat satu kandidat; ditemukan " & CStr(candidateCount) & "."
    End If
End Function

Private Sub ASRestoreState(ByVal doc As Document, ByVal previousLayer As Layer, _
                           ByRef layers() As Layer, _
                           ByRef wasEditable() As Boolean, ByVal previousUnit As cdrUnit, _
                           ByVal previousHUnit As cdrUnit, ByVal previousVUnit As cdrUnit)
    Dim i As Long
    For i = 1 To 3
        If Not layers(i) Is Nothing Then layers(i).Editable = wasEditable(i)
    Next i
    previousLayer.Activate
    doc.Unit = previousUnit
    doc.Rulers.HUnits = previousHUnit
    doc.Rulers.VUnits = previousVUnit
End Sub

Private Sub ASRecordCleanupIssue(ByRef issues As String, ByVal operation As String)
    If Err.Number <> 0 Then
        issues = issues & vbCrLf & "Pemulihan gagal (" & operation & "): " & Err.Description
    End If
    Err.Clear
End Sub
