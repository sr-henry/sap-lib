import win32com.client

## https://help.sap.com/viewer/b47d018c3b9b45e897faf66a6c0885a8/760.00/en-US

def create(profile: str) -> list:
    try:
        ## Instantiate SAP GUI application (creating the object)
        app = win32com.client.Dispatch("Sapgui.ScriptingCtrl.1") # GuiApplication Object
    except Exception as err:
        print(err.args[1])
        return []
    # Public Function OpenConnection( _
    #     ByVal Description As String, _
    #     Optional ByVal Sync As Variant, _
    #     Optional ByVal Raise As Variant _
    # ) As GuiConnection
    con = app.OpenConnection(profile, True, False) # GuiConnection Object
    
    if con is None:
        app = None
        print("Open Connection fail")
        return []

    print(con.Description, con.name, con.id, con.DisabledByServer)

    # This property is another name for the Children property
    session = con.Sessions(0) # GuiSession Object

    __multiple_logon(session)

    print(session.Info.User, session.Info.SystemName, session.Info.Client, session.name)

    return [app, con, session] # SAP Connection Data

def attach(profile: str) -> list: ## SAP Logon must be open
    try:
        ## Attach to a running instance of SAP GUI (getting the object)
        sap = win32com.client.GetObject("SAPGUI")
        ## Getting the scripting application
        app = sap.GetScriptingEngine # GuiApplication Object
    except Exception as err:
        print(err.args[1])
        return []

    # Public Function OpenConnection( _
    #     ByVal Description As String, _
    #     Optional ByVal Sync As Variant, _
    #     Optional ByVal Raise As Variant _
    # ) As GuiConnection
    con = app.OpenConnection(profile, True, False) # GuiConnection Object
    # In this case we're opening a new connection
    # however once we are getting a instance of SAP 
    # it's possible to get the a connecion that already exists
    # like this con.Children(0)
    if con is None:
        sap, app = None, None
        print("Open Connection fail")
        return []

    print(con.Description, con.name, con.id, con.DisabledByServer)

    # This property is another name for the Children property
    session = con.Sessions(0) # GuiSession Object

    __multiple_logon(session)

    print(session.Info.User, session.Info.SystemName, session.Info.Client, session.name)

    return [app, con, session] # SAP Connection Data

def __multiple_logon(session: win32com.client.CDispatch) -> None:
    while session.children.count > 1:
        try:
            session.findById("wnd[1]/usr/radMULTI_LOGON_OPT2").select()
            session.findById("wnd[1]/tbar[0]/btn[0]").press()
        except:
            session.ActiveWindow.sendVKey(0)

def close(sap_connection_data: list) -> None:
    sap_connection_data[1].CloseSession(sap_connection_data[2].id)
    sap_connection_data[1].CloseConnection()
    sap_connection_data[0] = None
    sap_connection_data[1] = None
    sap_connection_data[2] = None

def findByMapId(obj: win32com.client.CDispatch, elements: list, pos=0) -> win32com.client.CDispatch:
    if pos == len(elements):
        return obj
    try:
        for el in list(obj.children):
            if elements[pos] in el.id:
                el_id = findByMapId(el, elements, pos+1)
                if not el_id:
                    continue
                else:
                    return el_id
    except:
        return None

def debug(obj: win32com.client.CDispatch, depth = 0):
    try:
        for el in list(obj.children):
            print("\t"*depth, el.name, " -> ", el.text)
            debug(el, depth+1)
    except:
        return None

# When you extract the data from SAP in a local file 
# the default format is unique of SAP
# This method below try to convert this default format (unconvented)
# to a csv file, that can be read by any other tool
def unconverted_to_csv(input_file_name: str, skiprows=0, sep="|", output_encoding="utf-8") -> None:
    with open(input_file_name) as ifile, open(input_file_name + ".csv", "w", encoding=output_encoding) as ofile:
        data = ifile.readlines()
        rdata = ""
        for i,l in enumerate(data):
            if i < skiprows: continue
            if not l.strip(): continue
            if l[0] == '-': continue
            rdata += l[1:-1] + "\n"
        rdata = rdata.rstrip("\n")
        rdata = rdata.replace("|", sep)
        ofile.write(rdata)

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
