# SAP Gui Scripting API

Simple module to use sap gui scripting api in a easy way

### Documentation
 - https://help.sap.com/viewer/b47d018c3b9b45e897faf66a6c0885a8/760.00/en-US

### Requirements: 
- ``` pip install pywin32 ```

### How to use
```Python 3
    import sap

    sap_connection_data = sap.attach("System_Name")

    if sap_connection_data:

        application, connection, session = sap_connection_data

        # script here

        sap.close(sap_connection_data)
```

```VBA
    Dim sap As New SapGuiScripting
    If sap.Attach("System_Name") Then
        
        ' script here
        
    End If
```
