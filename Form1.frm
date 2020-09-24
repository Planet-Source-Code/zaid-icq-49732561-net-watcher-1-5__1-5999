VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H80000008&
   Caption         =   "NetWatcher 1.5"
   ClientHeight    =   4305
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   3165
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4305
   ScaleWidth      =   3165
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command2 
      Caption         =   "&Clear History"
      Height          =   375
      Left            =   120
      MaskColor       =   &H00FFFFFF&
      TabIndex        =   8
      Top             =   3000
      Width           =   2895
   End
   Begin VB.TextBox Text4 
      DataField       =   "TransactionType"
      DataSource      =   "Data1"
      Height          =   615
      Left            =   2040
      TabIndex        =   7
      Text            =   "Text3"
      Top             =   1440
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.TextBox Text3 
      DataField       =   "DateTime"
      DataSource      =   "Data1"
      Height          =   615
      Left            =   1680
      TabIndex        =   6
      Text            =   "Text3"
      Top             =   2040
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.Data Data1 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   "D:\Main Folder\0000000\Visual Basic Projects\NetWatch\netLog.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   495
      Left            =   3000
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "Log"
      Top             =   4200
      Visible         =   0   'False
      Width           =   1140
   End
   Begin VB.TextBox Text2 
      Height          =   495
      Left            =   960
      TabIndex        =   5
      Text            =   "Text2"
      Top             =   1560
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Hide (F9 to show)"
      Height          =   375
      Left            =   120
      MaskColor       =   &H00FFFFFF&
      TabIndex        =   4
      Top             =   3480
      Width           =   2895
   End
   Begin VB.TextBox Text1 
      Height          =   495
      Left            =   3120
      TabIndex        =   3
      Text            =   "Text1"
      Top             =   4200
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Timer Timer1 
      Interval        =   1
      Left            =   960
      Top             =   1320
   End
   Begin VB.ListBox List1 
      Height          =   2205
      Left            =   120
      TabIndex        =   0
      Top             =   720
      Width           =   2895
   End
   Begin VB.Label Label3 
      BackColor       =   &H000040C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Done"
      ForeColor       =   &H00FF0000&
      Height          =   255
      Left            =   120
      TabIndex        =   9
      Top             =   3960
      Width           =   2895
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackColor       =   &H80000012&
      Caption         =   "Label2"
      ForeColor       =   &H000000FF&
      Height          =   195
      Left            =   1320
      TabIndex        =   2
      Top             =   120
      Width           =   480
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H80000008&
      Caption         =   "Current Status:"
      ForeColor       =   &H00FF0000&
      Height          =   195
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   1050
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim results As Integer
Private Declare Function GetAsyncKeyState Lib "user32" (ByVal vKey As Long) As Integer

Private Declare Function RasEnumConnections Lib "RasApi32.dll" Alias "RasEnumConnectionsA" (lpRasCon As Any, lpcb As Long, lpcConnections As Long) As Long
Private Declare Function RasGetConnectStatus Lib "RasApi32.dll" Alias "RasGetConnectStatusA" (ByVal hRasCon As Long, lpStatus As Any) As Long
'
Private Const RAS95_MaxEntryName = 256
Private Const RAS95_MaxDeviceType = 16
Private Const RAS95_MaxDeviceName = 32
'
Private Type RASCONN95
    dwSize As Long
    hRasCon As Long
    szEntryName(RAS95_MaxEntryName) As Byte
    szDeviceType(RAS95_MaxDeviceType) As Byte
    szDeviceName(RAS95_MaxDeviceName) As Byte
End Type
'
Private Type RASCONNSTATUS95
    dwSize As Long
    RasConnState As Long
    dwError As Long
    szDeviceType(RAS95_MaxDeviceType) As Byte
    szDeviceName(RAS95_MaxDeviceName) As Byte
End Type
Dim laststausOn As Boolean
Dim connect As Boolean
Private Function IsConnected() As Boolean
Dim TRasCon(255) As RASCONN95
Dim lg As Long
Dim lpcon As Long
Dim RetVal As Long
Dim Tstatus As RASCONNSTATUS95
'
TRasCon(0).dwSize = 412
lg = 256 * TRasCon(0).dwSize

RetVal = RasEnumConnections(TRasCon(0), lg, lpcon)
If RetVal <> 0 Then
                    MsgBox "ERROR"
                    Exit Function
                    End If
'
Tstatus.dwSize = 160
RetVal = RasGetConnectStatus(TRasCon(0).hRasCon, Tstatus)
If Tstatus.RasConnState = &H2000 Then
                         IsConnected = True
                         connect = True
                         Else
                         IsConnected = False
                         connect = False
                         End If
 
End Function

Private Sub Command1_Click()
Form1.Hide
End Sub

Private Sub Command2_Click()
Label3.Caption = "Clearing history..."
Data1.RecordsetType = 0
Data1.Refresh
For i = 1 To Data1.Recordset.RecordCount
Data1.Recordset.MoveFirst
Data1.Recordset.Delete
Next
List1.Clear
Label3.Caption = "Done"
End Sub

Private Sub Form_Load()
Form1.Hide
App.TaskVisible = False
End Sub

Private Sub Text1_Change()
Data1.RecordsetType = 1
Data1.Refresh
If Text1.Text = True Then
Data1.Recordset.AddNew
Text4.Text = "Connected"
Text3.Text = Now
Data1.Recordset.Update
Data1.Refresh
List1.AddItem "Connected at " & Now
Label2.ForeColor = &HFF00&
ElseIf Text1.Text = "False" Then
Data1.Recordset.AddNew
Text4.Text = "Disconnected"
Text3.Text = Now
Data1.Recordset.Update
Data1.Refresh
Label2.ForeColor = &HFF&
List1.AddItem "Disconnected at " & Now
End If

End Sub

Private Sub Timer1_Timer()
IsConnected
If connect = True Then
Label2.Caption = "Connected"
Else
Label2.Caption = "Not Connected"
End If

Text1.Text = connect
 
For i = 1 To 255
results = 0
results = GetAsyncKeyState(i)
If results <> 0 And i = 120 Then
Form1.Show
End If
Next

End Sub
