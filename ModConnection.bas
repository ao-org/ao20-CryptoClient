Attribute VB_Name = "ModConnection"
Option Explicit

Public CH As CryptoHelper.CryptoHelper
Public CHinterface As CryptoHelper.ICryptoHelper
Public public_key() As Byte

Public Sub OpenSessionRequest()
    Dim arr(0 To 3) As Byte
    arr(0) = &H0
    arr(1) = &HAA
    arr(2) = &H0
    arr(3) = &H4
    Call Form1.Winsock1.SendData(arr)
    Form1.e_state = State.RequestOpenSession
End Sub

Public Sub AccountLoginRequest()
    Dim username As String
    Dim password As String
    Dim len_encrypted_password As Integer
    Dim len_encrypted_username As Integer
    Dim encrypted_username() As Byte
    Dim encrypted_password() As Byte
    
    Dim login_request() As Byte
    Dim packet_size As Integer
    Dim offset_login_request As Long
    Call AddtoRichTextBox("------------------------------------", 0, 255, 0, True)
    Call AddtoRichTextBox("AccountLoginRequest", 255, 255, 255, True)
    Call AddtoRichTextBox("------------------------------------", 0, 255, 0, True)
    username = Form1.txtUser.Text
    password = Form1.txtPass.Text
    
    encrypted_username = CHinterface.Encrypt(username, public_key)
    encrypted_password = CHinterface.Encrypt(password, public_key)
    
    Call AddtoRichTextBox("Username: " & Form1.ByteArr2String(encrypted_username), 255, 255, 255)
    Call AddtoRichTextBox("Password: " & Form1.ByteArr2String(encrypted_password), 255, 255, 255)
    
    
    ReDim login_request(1 To (2 + 2 + 2 + (UBound(encrypted_username) + 1) + 2 + (UBound(encrypted_password) + 1)))
    
    packet_size = UBound(login_request)
    
    login_request(1) = &HDE
    login_request(2) = &HAD
    'Siguientes 2 bytes indican tamaño total del paquete
    login_request(3) = Form1.hiByte(packet_size)
    login_request(4) = Form1.LoByte(packet_size)
    
    'Los siguientes 2 bytes son el SIZE_ENCRYPTED_USER
    
    len_encrypted_username = Len(Form1.ByteArr2String(encrypted_username))
    
    login_request(5) = Form1.hiByte(len_encrypted_username)
    login_request(6) = Form1.LoByte(len_encrypted_username)
    Call Form1.CopyBytes(encrypted_username, login_request, Len(Form1.ByteArr2String(encrypted_username)), 7)
    
    offset_login_request = 7 + UBound(encrypted_username)
    
    len_encrypted_password = Len(Form1.ByteArr2String(encrypted_password))
    
    login_request(offset_login_request + 1) = Form1.hiByte(len_encrypted_password)
    login_request(offset_login_request + 2) = Form1.LoByte(len_encrypted_password)
    
    Call Form1.CopyBytes(encrypted_password, login_request, Len(Form1.ByteArr2String(encrypted_password)), offset_login_request + 3)
    
    Call Form1.Winsock1.SendData(login_request)
    Form1.e_state = State.RequestAccountLogin
End Sub

Public Sub CreateAccountRequest()
    Dim arr() As Byte
    Dim packet_size As Integer
    Dim path As String, file As String
    path = App.path & "\character.txt"
    file = FileToString(path)
    Dim encrypted_account() As Byte
    
    encrypted_account = CHinterface.Encrypt(file, public_key)
    
    ReDim Preserve arr(1 To (2 + 2 + UBound(encrypted_account) + 1))
    
    arr(1) = &HBE
    arr(2) = &HEF
    
    packet_size = UBound(arr)
    
    arr(3) = Form1.hiByte(packet_size)
    arr(4) = Form1.LoByte(packet_size)
    
    Call Form1.CopyBytes(encrypted_account, arr, Len(Form1.ByteArr2String(encrypted_account)), 5)
    Call Form1.Winsock1.SendData(arr)
    Form1.e_state = State.RequestAccountCreate
End Sub
Public Sub connectToLoginServer()
    Form1.Winsock1.RemoteHost = "194.113.72.86"
    Form1.Winsock1.RemotePort = "4000"
    Form1.Winsock1.Connect
End Sub

Public Sub HandleOpenSession(ByVal BytesTotal As Long)
    Call AddtoRichTextBox("------------------------------------", 0, 255, 0, True)
    Call AddtoRichTextBox("HandleOpenSession", 255, 255, 255, True)
    Call AddtoRichTextBox("------------------------------------", 0, 255, 0, True)
    Dim strData As String
    Form1.Winsock1.PeekData strData, vbString, BytesTotal
    Call AddtoRichTextBox("Bytes total: " & strData, 255, 255, 255, False)
    
    Form1.Winsock1.GetData strData, vbString, 2
    Call AddtoRichTextBox("Id: " & strData, 255, 255, 255, False)
    
    Form1.Winsock1.GetData strData, vbString, 2
    
    Dim encrypted_token() As Byte
    Dim secret_key_byte() As Byte
    
    Form1.Winsock1.GetData encrypted_token, 64
            
    Call Form1.Str2ByteArr("pablomarquezARG1", secret_key_byte)
    Dim decrypted_session_token As String
     
    decrypted_session_token = CHinterface.Decrypt(encrypted_token, secret_key_byte)
    Call AddtoRichTextBox("Decripted_session_token: " & decrypted_session_token, 255, 255, 255, False)
        
    public_key = Mid(decrypted_session_token, 1, 16)
    
    Call AddtoRichTextBox("Public key:" & CStr(public_key), 255, 255, 255, False)
    
    Form1.Str2ByteArr decrypted_session_token, public_key, 16
    Form1.e_state = State.SessionOpen
    
End Sub

Public Sub HandleAccountLogin(ByVal BytesTotal As Long)

    Call AddtoRichTextBox("------------------------------------", 0, 255, 0, True)
    Call AddtoRichTextBox("HandleRequestAccountLogin", 255, 255, 255, True)
    Call AddtoRichTextBox("------------------------------------", 0, 255, 0, True)
    Dim data() As Byte
    
    Form1.Winsock1.PeekData data, vbByte, BytesTotal
    
    Form1.Winsock1.GetData data, vbByte, 2
    
    If data(0) = &HAF And data(1) = &HA1 Then
        Call AddtoRichTextBox("LOGIN-OK", 0, 255, 0, True)
    Else
       Call AddtoRichTextBox("ERROR", 255, 0, 0, True)
    End If
        
    Call AddtoRichTextBox(Form1.ByteArrayToHex(data), 255, 255, 255)
    Form1.Winsock1.GetData data, vbByte, 2
    Form1.e_state = State.SessionOpen
End Sub

Public Sub HandleAccountCreate(ByVal BytesTotal As Long)

    Call AddtoRichTextBox("------------------------------------", 0, 255, 0, True)
    Call AddtoRichTextBox("HandleAccountCreate", 255, 255, 255, True)
    Call AddtoRichTextBox("------------------------------------", 0, 255, 0, True)
    Dim data() As Byte
    
    Form1.Winsock1.PeekData data, vbByte, BytesTotal
    Call AddtoRichTextBox(Form1.ByteArrayToHex(data), 255, 255, 255)
    Form1.Winsock1.GetData data, vbByte, 2
    
    If data(0) = &HAF And data(1) = &HA1 Then
        Call AddtoRichTextBox("ACCOUNT-CREATED-OK", 0, 255, 0, True)
    Else
       Call AddtoRichTextBox("ERROR", 255, 0, 0, True)
    End If
        
    Debug.Print Form1.ByteArrayToHex(data)
    Form1.Winsock1.GetData data, vbByte, 2
    Form1.e_state = State.SessionOpen
End Sub


Sub AddtoRichTextBox(ByVal Text As String, Optional ByVal red As Integer = -1, Optional ByVal green As Integer, Optional ByVal blue As Integer, Optional ByVal bold As Boolean = False, Optional ByVal italic As Boolean = False, Optional ByVal bCrLf As Boolean = False)
    '******************************************
    'HarThaoS: Martín Trionfetti 20-11-2021
    '******************************************
    With Form1.rtbConsole
    
        If Len(.Text) > 20000 Then
            .Text = vbNullString
            .SelStart = InStr(1, .Text, vbCrLf) + 1
            .SelLength = Len(.Text) - .SelStart + 2
            .TextRTF = .SelRTF
        End If
        
        .SelStart = Len(.Text)
        .SelLength = 0
        .SelBold = bold
        .SelItalic = italic
        
        If Not red = -1 Then .SelColor = RGB(red, green, blue)
        bCrLf = True
        
        If bCrLf And Len(.Text) > 0 Then Text = vbCrLf & Text
        .SelText = Text
        
    End With
    
End Sub

Function FileToString(strFilename As String) As String
  Open strFilename For Input As #1
    FileToString = StrConv(InputB(LOF(1), 1), vbUnicode)
  Close #1
End Function
