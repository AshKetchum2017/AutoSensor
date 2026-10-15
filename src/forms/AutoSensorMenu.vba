Option Explicit

Private mProcessing As Boolean
Private Const SENSOR_DEFAULT_SIZE As String = "7"
Private Const SENSOR_SIZE_HINT As String = "7 mm"
Private mHintReady As Boolean
Private mUpdatingHint As Boolean
Private mSensorHintActive As Boolean
Private mSensorSizeFocused As Boolean
Private mSensorInputColor As Long

Private Sub UserForm_Initialize()
    mSensorInputColor = txbSensorSize.ForeColor
    mHintReady = True
    If Not optSensor.Value And Not optWhiteSpace.Value Then optSensor.Value = True
    UpdateSizeInput
End Sub

Private Sub cmdClose_Click()
    If Not mProcessing Then Unload Me
End Sub

Private Sub cmdProcess_Click()
    If mProcessing Then Exit Sub
    If optSensor.Value = optWhiteSpace.Value Then
        MsgBox "Pilih salah satu mode: Sensor atau WhiteSpace.", vbExclamation, "AutoSensor"
        Exit Sub
    End If

    On Error GoTo ProcessFailed
    mProcessing = True
    cmdProcess.Enabled = False
    ASGenerate optWhiteSpace.Value, SensorSizeInput()

ProcessFinished:
    mProcessing = False
    cmdProcess.Enabled = True
    UpdateSizeInput
    Exit Sub

ProcessFailed:
    MsgBox "AutoSensor gagal: " & Err.Description, vbExclamation, "AutoSensor"
    Resume ProcessFinished
End Sub

Private Sub optSensor_Click()
    UpdateSizeInput
End Sub

Private Sub optWhiteSpace_Click()
    UpdateSizeInput
End Sub

Private Sub txbSensorSize_Change()
    If Not mHintReady Or mUpdatingHint Then Exit Sub
    mSensorHintActive = False
    txbSensorSize.ForeColor = mSensorInputColor
    ShowSensorSizeHint
End Sub

Private Sub txbSensorSize_Enter()
    If Not mHintReady Then Exit Sub
    mSensorSizeFocused = True
    mUpdatingHint = True
    If mSensorHintActive Then txbSensorSize.Value = vbNullString
    mSensorHintActive = False
    txbSensorSize.ForeColor = mSensorInputColor
    mUpdatingHint = False
End Sub

Private Sub txbSensorSize_Exit(ByVal Cancel As MSForms.ReturnBoolean)
    mSensorSizeFocused = False
    ShowSensorSizeHint
End Sub

Private Sub UpdateSizeInput()
    txbSensorSize.Enabled = optSensor.Value And Not optWhiteSpace.Value
    If Not txbSensorSize.Enabled Then mSensorSizeFocused = False
    ShowSensorSizeHint
End Sub

Private Sub ShowSensorSizeHint()
    If Not mHintReady Or mUpdatingHint Or mSensorSizeFocused Then Exit Sub
    If Len(Trim$(CStr(txbSensorSize.Value))) > 0 Then Exit Sub
    mUpdatingHint = True
    mSensorHintActive = True
    txbSensorSize.ForeColor = RGB(128, 128, 128)
    txbSensorSize.Value = SENSOR_SIZE_HINT
    mUpdatingHint = False
End Sub

Private Function SensorSizeInput() As String
    If mSensorHintActive Or Len(Trim$(CStr(txbSensorSize.Value))) = 0 Then
        SensorSizeInput = SENSOR_DEFAULT_SIZE
    Else
        SensorSizeInput = CStr(txbSensorSize.Value)
    End If
End Function
