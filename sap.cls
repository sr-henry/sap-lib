VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "SapGuiScripting"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

' Import SAP GUI Scripting API (sapfewse.ocx)

Option Explicit

Private Declare PtrSafe Function CoRegisterMessageFilter Lib "ole32.dll" (ByVal IFilterIn As Long, ByRef PreviousFilter As Long) As Long

Private gui As Object
Private app As GuiApplication
Private con As GuiConnection
Public session As GuiSession

Private Sub Class_Terminate()
    If Not C1ose() Then
        Debug.Print "SAP Connection was not closed"
    End If
End Sub

Public Function Attach( _
    profile As String, _
    Optional ByVal Inplace As Boolean = False _
) As Boolean
    
    ' Attach to a running instance of SAP GUI (getting the object)
    On Error Resume Next
    Set gui = GetObject("SAPGUI")
    On Error GoTo 0
    
    If gui Is Nothing Then
        Debug.Print "[ ! ] SAP Logon instance was not found"
        Attach = False
        Exit Function
    End If
    
    ' Getting the scripting application
    Set app = gui.GetScriptingEngine
    
    Attach = CreateConnection(profile, Inplace)
    
End Function

Private Function CreateConnection( _
    profile As String, _
    Inplace As Boolean _
) As Boolean

    ' To create a new SAP GUI instance placed within your application
    Dim inplace_suffix As String
    If Inplace Then inplace_suffix = "/INPLACE" Else inplace_suffix = ""
    
    On Error Resume Next
    Set con = app.OpenConnection(profile & inplace_suffix, True, False)
    On Error GoTo 0
    
    If con Is Nothing Then
        Debug.Print "[ ! ] Open Connection fail"
        Set app = Nothing
        CreateConnection = False
        Exit Function
    End If
    
    Debug.Print con.Description & "    " & con.Name
    
    Set session = con.Sessions(0)
    
    Call MultipleLogon(session)

    session.ActiveWindow.Maximize

    Debug.Print session.info.User & "    " & session.info.SystemName & "    " & session.info.Client & "    " & session.Name

    Call CoRegisterMessageFilter(0, 0)

    CreateConnection = True

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
    con.CloseSession session.ID
    con.CloseConnection
    Set app = Nothing
    Set con = Nothing
    Set session = Nothing
    C1ose = True
eh:
End Function
