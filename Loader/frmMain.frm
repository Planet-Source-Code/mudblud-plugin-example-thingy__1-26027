VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmMain 
   Caption         =   "Plugin Loader"
   ClientHeight    =   2070
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5805
   LinkTopic       =   "Form1"
   ScaleHeight     =   2070
   ScaleWidth      =   5805
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame3 
      Caption         =   "RegSvr"
      Height          =   2055
      Left            =   4920
      TabIndex        =   6
      Top             =   0
      Width           =   855
      Begin VB.CommandButton Command2 
         Caption         =   "UnReg"
         Height          =   1695
         Left            =   480
         TabIndex        =   7
         Top             =   240
         Width           =   255
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Reg"
         Height          =   1695
         Left            =   120
         TabIndex        =   8
         Top             =   240
         Width           =   255
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Plugins"
      Height          =   2055
      Left            =   3360
      TabIndex        =   1
      Top             =   0
      Width           =   1455
      Begin VB.CommandButton Command3 
         Caption         =   "Settings"
         Enabled         =   0   'False
         Height          =   255
         Left            =   120
         TabIndex        =   9
         Top             =   1680
         Width           =   1215
      End
      Begin VB.CommandButton cmdUnload 
         Caption         =   "Unload"
         Height          =   255
         Left            =   120
         TabIndex        =   4
         Top             =   1440
         Width           =   1215
      End
      Begin VB.CommandButton cmdLoad 
         Caption         =   "Load"
         Height          =   255
         Left            =   120
         TabIndex        =   3
         Top             =   1200
         Width           =   1215
      End
      Begin VB.ListBox lstPlugins 
         Height          =   840
         Left            =   120
         TabIndex        =   2
         Top             =   240
         Width           =   1215
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Message"
      Height          =   2055
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   3255
      Begin VB.TextBox txtMessage 
         Height          =   1695
         Left            =   120
         MultiLine       =   -1  'True
         TabIndex        =   5
         Top             =   240
         Width           =   3015
      End
      Begin MSComDlg.CommonDialog CD1 
         Left            =   2970
         Top             =   120
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdLoad_Click()
Dim Tryed2Reg As Boolean
On Error Resume Next
CD1.InitDir = App.Path
CD1.Filter = "DLL Plugins (*.dll)|*.dll|"
CD1.ShowOpen
If CD1.FileName = "" Then Exit Sub 'if they clicked cancel dont try and do ne thing
'---------
Dim FN As String
FN = Mid$(CD1.FileName, InStrRev(CD1.FileName, "\") + 1) 'gets the filename from the path e.g. test.dll
FN = Mid$(FN, 1, Len(FN) - 4) 'removes the .dll part
'---------
TryCreate:
Dim tmpObject
Set tmpObject = CreateObject(FN & ".Plugin") 'try and create object
If tmpObject Is Nothing Then 'didn't create
If Tryed2Reg = True Then ' its been tryed b4 but it still wont work, my guess is it aint a plugin.
MsgBox "This plugin couldn't be loaded."
Else 'ok, lets try and register it.
ret = RegisterServer(Me.hWnd, CD1.FileName, True)
Tryed2Reg = True
GoTo TryCreate 'go back and try create the object again
End If
Else ' Plugin was loaded.
lstPlugins.AddItem FN
End If
Set tmpObject = Nothing
End Sub

Private Sub cmdUnload_Click()
If lstPlugins.ListIndex < 0 Then Exit Sub ' dont try if a item aint selected
lstPlugins.RemoveItem lstPlugins.ListIndex
txtMessage.Text = ""
End Sub

Private Sub Command1_Click()
CD1.InitDir = App.Path
CD1.Filter = "DLL Plugins (*.dll)|*.dll|"
CD1.ShowOpen
If CD1.FileName = "" Then Exit Sub
RegisterServer Me.hWnd, CD1.FileName, True
If reg = ERROR_SUCCESS Then
MsgBox "Registered"
Else
MsgBox "Error"
End If
End Sub

Private Sub Command2_Click()
CD1.InitDir = App.Path
CD1.Filter = "DLL Plugins (*.dll)|*.dll|"
CD1.ShowOpen
If CD1.FileName = "" Then Exit Sub
ret = RegisterServer(Me.hWnd, CD1.FileName, False)
If reg = ERROR_SUCCESS Then
MsgBox "Unregistered"
Else
MsgBox "Error"
End If
End Sub

Private Sub Command3_Click()
If lstPlugins.ListIndex < 0 Then Exit Sub ' dont try if a item aint selected
Dim tmpObject As Object
Set tmpObject = CreateObject(lstPlugins.List(lstPlugins.ListIndex) & ".Plugin")
tmpObject.ShowSettings
Set tmpObject = Nothing
End Sub

Private Sub lstPlugins_Click()
If lstPlugins.ListIndex < 0 Then Command3.Enabled = False: Exit Sub ' dont try if a item aint selected
Dim tmpObject As Object
Set tmpObject = CreateObject(lstPlugins.List(lstPlugins.ListIndex) & ".Plugin")
txtMessage.Text = tmpObject.message
Set tmpObject = Nothing
Command3.Enabled = True
End Sub
