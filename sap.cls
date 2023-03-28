VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "SapGuiScripting"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' =========================================================================================='
' Author: Henrique Campiotti (ehnraem)                                                      '
' Description: Simple class module to use SAP GUI Scripting API in a easy way               '
' Documentation: https://help.sap.com/viewer/b47d018c3b9b45e897faf66a6c0885a8/760.00/en-US  '
' =========================================================================================='

Option Explicit

Private Declare PtrSafe Function CoRegisterMessageFilter Lib "ole32.dll" (ByVal IFilterIn As Long, ByRef PreviousFilter As Long) As Long

Private Const mTab = "    "

Private gui As Object
Private app As GuiApplication
Private con As GuiConnection
Public session As GuiSession

Private Sub Class_Terminate()
    If Not C1ose() Then Application.StatusBar = "SAP Connection was not closed"
End Sub

Public Function Attach( _
    profile As String, _
    Optional ByVal inplace As Boolean = False, _
    Optional ByVal usr As String, _
    Optional ByVal pwd As String _
    ) As Boolean
    
    ' Attach to a running instance of SAP GUI (getting the object)
    On Error Resume Next
    Set gui = GetObject("SAPGUI")
    On Error GoTo 0
    
    If gui Is Nothing Then
        Application.StatusBar = "[ ! ] SAP Logon instance was not found"
        Attach = False
        Exit Function
    End If
    
    ' Getting the scripting application
    Set app = gui.GetScriptingEngine
    
    Attach = CreateConnection(profile, inplace, usr, pwd)
    
End Function

Private Function CreateConnection( _
    profile As String, _
    inplace As Boolean, _
    Optional usr As String, _
    Optional pwd As String _
    ) As Boolean

    ' To create a new SAP GUI instance placed within your application
    Dim inplSuffix$: If inplace Then inplSuffix = "/INPLACE"
    
    On Error Resume Next
    Set con = app.OpenConnection(profile & inplSuffix, True, False)
    On Error GoTo 0
    
    If con Is Nothing Then
        Application.StatusBar = "[ ! ] Open Connection fail"
        Set app = Nothing
        Exit Function
    End If
    
    Set session = con.Sessions(0)
    
    If Not Logon(usr, pwd) Then Exit Function
    
    Call MultipleLogon(session)

    session.ActiveWindow.Maximize

    Application.StatusBar = session.info.User & mTab & session.info.SystemName & mTab & _
                            session.info.Client & mTab & session.Name

    Call CoRegisterMessageFilter(0, 0)

    CreateConnection = True

End Function

Private Function Logon(usr As String, pwd As String) As Boolean
    
    On Error Resume Next
    Dim usrField As GuiTextField, pwdField As GuiPasswordField
    With session.ActiveWindow
        Set usrField = .FindById("usr/txtRSYST-BNAME")
        Set pwdField = .FindById("usr/pwdRSYST-BCODE")
    End With
    On Error GoTo 0
    
    If usrField Is Nothing And pwdField Is Nothing Then
        Application.StatusBar = "[ + ] Single sign-on active"
        Logon = True
        Exit Function
    End If
    
    usrField.text = usr
    pwdField.text = pwd
    session.ActiveWindow.SendVKey 0
    
    Dim sts As GuiStatusbar
    Set sts = session.ActiveWindow.FindById("sbar")
    
    If sts.MessageType = "E" Then
        Application.StatusBar = sts.text
        Exit Function
    End If
    
    Logon = True
    
End Function

Private Sub MultipleLogon(session As GuiSession)
    While session.Children.Count > 1:
        On Error Resume Next
        session.FindById("wnd[1]/usr/radMULTI_LOGON_OPT2").Select
        session.ActiveWindow.SendVKey 0
        On Error GoTo 0
    Wend
End Sub

Private Function C1ose() As Boolean
    On Error GoTo eh:
    If con Is Nothing Or gui Is Nothing Then GoTo dn
    con.CloseSession session.id
    con.CloseConnection
    Set app = Nothing
    Set con = Nothing
    Set session = Nothing
dn:
    C1ose = True
eh:
End Function
