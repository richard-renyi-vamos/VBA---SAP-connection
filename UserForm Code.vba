Option Explicit

' Login Button Click Event
Private Sub btnLogin_Click()
    If ConnectToSAP Then
        MsgBox "Connected to SAP Successfully!", vbInformation, "SAP Connection"
    Else
        MsgBox "Failed to connect to SAP. Make sure SAP is open.", vbCritical, "Error"
    End If
End Sub

' Execute Transaction Button Click Event
Private Sub btnExecute_Click()
    Dim TCode As String
    TCode = txtTCode.Text
    
    If TCode = "" Then
        MsgBox "Please enter a transaction code.", vbExclamation, "Input Required"
        Exit Sub
    End If
    
    ExecuteSAPTransaction TCode
End Sub
