VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "SapAnalysisOffice"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private xl As excel.Application
Private m_client As String
Private m_user As String

Private Sub class_initialize()
    Set xl = ThisWorkbook.Application
    m_client = "000"
    m_user = Environ$("username")
End Sub

Public Property Set ExcelApplication(ByRef xlApp As excel.Application)
    Set xl = xlApp
End Property

Public Property Let Client(ByVal clnt As String)
    m_client = clnt
End Property

Public Property Let User(ByVal usr As String)
    m_user = usr
End Property

Public Function Refresh(dataSrc As String, prompts As Dictionary) As Boolean
   
    If Not ActivateAnalysisAddin(xl) Then
        Debug.Print "Analysis addin not found!"
        Exit Function
    End If

    Dim lResult As Long
    
    lResult = xl.Run("SAPLogon", dataSrc, m_client, m_user) ' SSO
    If Not CBool(lResult) Then
        Debug.Print "SAPLogon Error"
        Exit Function
    End If

    lResult = xl.Run("SAPExecuteCommand", "RefreshData", dataSrc)
    If Not CBool(lResult) Then
        Debug.Print "RefreshData Error"
        Exit Function
    End If
    
    On Error GoTo eh
    Call xl.Run("SAPExecuteCommand", "PauseVariableSubmit", "On")
    Call xl.Run("SAPSetRefreshBehaviour", "Off")
    
    Dim key As Variant
    For Each key In prompts
        Call xl.Run("SAPSetVariable", key, prompts(key), "INPUT_STRING", dataSrc)
    Next key
    
    Call xl.Run("SAPExecuteCommand", "PauseVariableSubmit", "Off")
    Call xl.Run("SAPSetRefreshBehaviour", "On")
    
    Refresh = True
    Exit Function
eh:
    Debug.Print Err.Description, Err.source
End Function

Private Function ActivateAnalysisAddin(excel As excel.Application) As Boolean
    
    Dim addin As COMAddIn
    For Each addin In excel.Application.COMAddIns
        If addin.progID = "SapExcelAddIn" Then
            addin.Connect = False
            addin.Connect = True
            ActivateAnalysisAddin = True
            Exit For
        End If
    Next
    
End Function

