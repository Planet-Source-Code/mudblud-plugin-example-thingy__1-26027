VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Plugin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Function Message() As String
Message = GetSetting("TestPlugin", "Message", "Message", "No message stored" & vbCrLf & "Click settings to put in a message.")
End Function

Sub ShowSettings()
frmSettings.Show vbModal
End Sub
