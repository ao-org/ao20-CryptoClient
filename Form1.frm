VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "richtx32.ocx"
Begin VB.Form Form1 
   AutoRedraw      =   -1  'True
   BackColor       =   &H80000001&
   BorderStyle     =   5  'Sizable ToolWindow
   Caption         =   "Encryptor"
   ClientHeight    =   2640
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   11415
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2640
   ScaleWidth      =   11415
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtPass 
      Alignment       =   2  'Center
      BackColor       =   &H80000006&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000C000&
      Height          =   285
      IMEMode         =   3  'DISABLE
      Left            =   120
      PasswordChar    =   "*"
      TabIndex        =   7
      Text            =   "Pablo17"
      Top             =   960
      Width           =   2775
   End
   Begin VB.TextBox txtUser 
      Alignment       =   2  'Center
      BackColor       =   &H80000006&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000C000&
      Height          =   285
      Left            =   120
      TabIndex        =   5
      Text            =   "roger3"
      Top             =   360
      Width           =   2775
   End
   Begin RichTextLib.RichTextBox rtbConsole 
      Height          =   2055
      Left            =   3000
      TabIndex        =   4
      Top             =   120
      Width           =   8295
      _ExtentX        =   14631
      _ExtentY        =   3625
      _Version        =   393217
      BackColor       =   -2147483647
      BorderStyle     =   0
      ScrollBars      =   2
      TextRTF         =   $"Form1.frx":0000
   End
   Begin VB.Timer timerConnection 
      Enabled         =   0   'False
      Interval        =   500
      Left            =   4440
      Top             =   120
   End
   Begin VB.CommandButton cmdCreateAccount 
      Caption         =   "Create Account"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   1560
      Width           =   2775
   End
   Begin VB.CommandButton cmdLoginAccount 
      Caption         =   "Login"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   1320
      Width           =   2775
   End
   Begin MSWinsockLib.Winsock Winsock1 
      Left            =   0
      Top             =   0
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Username"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000B&
      Height          =   195
      Left            =   1125
      TabIndex        =   6
      Top             =   120
      Width           =   870
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Password"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000B&
      Height          =   195
      Left            =   1140
      TabIndex        =   3
      Top             =   720
      Width           =   840
   End
   Begin VB.Label lblStatus 
      Alignment       =   2  'Center
      BackColor       =   &H80000008&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Not connected"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   375
      Left            =   480
      TabIndex        =   0
      Top             =   1920
      Width           =   2055
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public Enum State
    Idle = 0
    RequestOpenSession
    SessionOpen
    RequestAccountLogin
    RequestAccountCreate
End Enum

Public e_state As State

Private Sub cmdCreateAccount_Click()
    If e_state = State.SessionOpen Then
        e_state = State.RequestAccountCreate
        Call CreateAccountRequest
    End If
End Sub

Private Sub cmdLoginAccount_Click()
     If e_state = State.SessionOpen Then
        e_state = State.RequestAccountLogin
        Call AccountLoginRequest
    End If
End Sub

Private Sub Form_Load()
    Set CH = New CryptoHelper.CryptoHelper
    Set CHinterface = CH
    e_state = State.Idle
    
    If Winsock1.State <> 7 Then
        Call connectToLoginServer
        Form1.timerConnection.Enabled = True
        If Winsock1.State <> 7 Then
            lblStatus.ForeColor = vbYellow
            lblStatus.Caption = "Connecting..."
        End If
    End If
    
End Sub

Private Sub timerConnection_Timer()
    If Winsock1.State = 7 Then
        timerConnection.Enabled = False
        lblStatus.Caption = "Connected"
        lblStatus.ForeColor = vbGreen
        Call OpenSessionRequest
    End If
End Sub

Private Sub Winsock1_DataArrival(ByVal BytesTotal As Long)
    Select Case e_state
        Case State.RequestOpenSession
            Call HandleOpenSession(BytesTotal)
        Case State.RequestAccountLogin
            Call HandleAccountLogin(BytesTotal)
        Case State.RequestAccountCreate
            Call HandleAccountCreate(BytesTotal)
    End Select
End Sub

'HarThaoS: Convierto el str en arr() bytes
Public Function Str2ByteArr(ByVal str As String, ByRef arr() As Byte, Optional ByVal length As Long = 0)
    Dim i As Long
    Dim asd As String
    If length = 0 Then
        ReDim arr(0 To (Len(str) - 1))
        For i = 0 To (Len(str) - 1)
            arr(i) = Asc(Mid(str, i + 1, 1))
        Next i
    Else
        ReDim arr(0 To (length - 1)) As Byte
        For i = 0 To (length - 1)
            arr(i) = Asc(Mid(str, i + 1, 1))
        Next i
    End If
    
End Function

Public Function ByteArr2String(ByRef arr() As Byte) As String
    
    Dim str As String
    Dim i As Long
    For i = 0 To UBound(arr)
        str = str + Chr(arr(i))
    Next i
    
    ByteArr2String = str
    
End Function

Public Function hiByte(ByVal w As Integer) As Byte
    Dim hi As Integer
    If w And &H8000 Then hi = &H4000
    
    hiByte = (w And &H7FFE) \ 256
    hiByte = (hiByte Or (hi \ 128))
    
End Function

Public Function LoByte(w As Integer) As Byte
 LoByte = w And &HFF
End Function

Public Function MakeInt(ByVal LoByte As Byte, _
   ByVal hiByte As Byte) As Integer

MakeInt = ((hiByte * &H100) + LoByte)

End Function

Public Function CopyBytes(ByRef src() As Byte, ByRef dst() As Byte, ByVal size As Long, Optional ByVal offset As Long = 0)
    Dim i As Long
    
    For i = 0 To (size - 1)
        dst(i + offset) = src(i)
    Next i
    
End Function

Public Function ByteArrayToHex(ByRef ByteArray() As Byte) As String
    Dim l As Long, strRet As String
    
    For l = LBound(ByteArray) To UBound(ByteArray)
        strRet = strRet & Hex$(ByteArray(l)) & " "
    Next l
    
    'Remove last space at end.
    ByteArrayToHex = Left$(strRet, Len(strRet) - 1)
End Function

