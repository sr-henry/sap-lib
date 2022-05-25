import win32com.client as __win32
import win32clipboard as __clip

## Requirements: pip install pywin32
## Documentation: https://help.sap.com/viewer/b47d018c3b9b45e897faf66a6c0885a8/760.00/en-US

def create(profile: str, inplace: bool = False) -> list:
    try:
        ## Instantiate SAP GUI application (creating the object)
        app = __win32.Dispatch("Sapgui.ScriptingCtrl.1") # GuiApplication Object
    except Exception as err:
        print(f"[!] Unable to create SAP GUI instance for scripting: {err.args[1]}")
        return []

    return __create_connection(profile, [app], inplace)


def attach(profile: str, inplace: bool = False) -> list:
    try:
        ## Attach to a running instance of SAP GUI (getting the object)
        sap = __win32.GetObject("SAPGUI")
    except Exception as err:
        print(f"[!] SAP Logon instance was not found: {err.args[1]}")
        return []

    ## Getting the scripting application
    app = sap.GetScriptingEngine # GuiApplication Object
    if not isinstance(app, __win32.CDispatch):
        sap = None
        return []

    return __create_connection(profile, [app], inplace)
    

def __create_connection(profile: str, lapp: list, inplace: bool) -> list:
    ## To create a new SAP GUI instance placed within your application
    profile += "/INPLACE" if inplace else "" 

    # Public Function OpenConnection( _
    #     ByVal Description As String, _
    #     Optional ByVal Sync As Variant, _
    #     Optional ByVal Raise As Variant _
    # ) As GuiConnection
    con = lapp[0].OpenConnection(profile, True, False) # GuiConnection Object

    # In this case we're opening a new connection
    # however once we are getting a instance of SAP 
    # it's possible to get the a connecion that already exists
    # like this con.Children(0)
    if con is None:
        sap, lapp[0] = None, None
        print("Open Connection fail")
        return []

    print(f"{con.Description} {con.name}")

    # This property is another name for the Children property
    session = con.Sessions(0) # GuiSession Object

    __multiple_logon(session)

    print(f"{session.Info.User} {session.Info.SystemName} {session.Info.Client} {session.name}")

    return [lapp[0], con, session] # SAP Connection Data


def __multiple_logon(session: object) -> None:
    while session.children.count > 1:
        try:
            session.FindById("wnd[1]/usr/radMULTI_LOGON_OPT2").select()
            session.FindById("wnd[1]/tbar[0]/btn[0]").press()
        except:
            session.ActiveWindow.sendVKey(vKeys["Enter"])


def close(sap_connection_data: list) -> bool:
    try:
        sap_connection_data[1].CloseSession(sap_connection_data[2].id)
        sap_connection_data[1].CloseConnection()
        sap_connection_data[0] = None
        sap_connection_data[1] = None
        sap_connection_data[2] = None
        return True
    except Exception as err:
        print(f"[!] SAP Connection was not closed: {err.args[1]}")
        return False

def list_to_clipboard(data: list) -> None:
    text = "\r\n".join(str(e) for e in data)
    __clip.OpenClipboard()
    __clip.EmptyClipboard()
    __clip.SetClipboardText(text)
    __clip.CloseClipboard()


vKeys = {
    "Enter": 0,
    "F1": 1,
    "F2": 2,
    "F3": 3,
    "F4": 4,
    "F5": 5,
    "F6": 6,
    "F7": 7,
    "F8": 8,
    "F9": 9,
    "F10": 10,
    "Ctrl+S": 11,
    "F12": 12,
    "Shift+F1": 13,
    "Shift+F2": 14,
    "Shift+F3": 15,
    "Shift+F4": 16,
    "Shift+F5": 17,
    "Shift+F6": 18,
    "Shift+F7": 19,
    "Shift+F8": 20,
    "Shift+F9": 21,
    "Shift+Ctrl+0": 22,
    "Shift+F11": 23,
    "Shift+F12": 24,
    "Ctrl+F1": 25,
    "Ctrl+F2": 26,
    "Ctrl+F3": 27,
    "Ctrl+F4": 28,
    "Ctrl+F5": 29,
    "Ctrl+F6": 30,
    "Ctrl+F7": 31,
    "Ctrl+F8": 32,
    "Ctrl+F9": 33,
    "Ctrl+F10": 34,
    "Ctrl+F11": 35,
    "Ctrl+F12": 36,
    "Ctrl+Shift+F1": 37,
    "Ctrl+Shift+F2": 38,
    "Ctrl+Shift+F3": 39,
    "Ctrl+Shift+F4": 40,
    "Ctrl+Shift+F5": 41,
    "Ctrl+Shift+F6": 42,
    "Ctrl+Shift+F7": 43,
    "Ctrl+Shift+F8": 44,
    "Ctrl+Shift+F9": 45,
    "Ctrl+Shift+F10": 46,
    "Ctrl+Shift+F11": 47,
    "Ctrl+Shift+F12": 48
}