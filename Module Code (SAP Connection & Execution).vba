Option Explicit

Public SapGui As Object
Public Application As Object
Public Connection As Object
Public Session As Object

' Function to Connect to SAP
Function ConnectToSAP() As Boolean
    On Error Resume Next
    Set SapGui = GetObject("SAPGUI")
    Set Application = SapGui.GetScriptingEngine
    Set Connection = Application.Children(0)
    Set Session = Connection.Children(0)
    
    If Err.Number <> 0 Then
        MsgBox "Error connecting to SAP. Please ensure SAP is open.", vbCritical, "SAP Connection Error"
        ConnectToSAP = False
        Exit Function
    End If
    
    ConnectToSAP = True
End Function

' Function to Execute a Transaction
Sub ExecuteSAPTransaction(TCode As String)
    If Not ConnectToSAP Then Exit Sub
    Session.findById("wnd[0]/tbar[0]/okcd").Text = TCode
    Session.findById("wnd[0]").sendVKey 0
    MsgBox "Transaction " & TCode & " executed successfully.", vbInformation, "SAP Execution"
End Sub
