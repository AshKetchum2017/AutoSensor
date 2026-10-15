Option Explicit

Private mProcessing As Boolean

Private Sub UserForm_Initialize()
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
    ASGenerate optWhiteSpace.Value, CStr(txbSensorSize.Value)

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
    ' Ukuran divalidasi saat Process, bukan ketika pengguna masih mengetik.
End Sub

Private Sub UpdateSizeInput()
    txbSensorSize.Enabled = optSensor.Value And Not optWhiteSpace.Value
End Sub
