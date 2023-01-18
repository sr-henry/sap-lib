# SAP Automation Modules

Simple modules to use sap gui scripting api and sap analysis for office in a easy way 😀

### Documentations

 - [SAP GUI Scripting API](https://help.sap.com/viewer/b47d018c3b9b45e897faf66a6c0885a8/760.00/en-US)
 - [SAP Analysis for Microsoft Office](https://help.sap.com/docs/SAP_BUSINESSOBJECTS_ANALYSIS_OFFICE/ca9c58444d64420d99d6c136a3207632/ebf198667aa54740b9049d9da804a901.html?version=2.8.8.0)

### Gui Scripting Module
```Python 3
    # pip install pywin32
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

### Analysis for Office Module

```VBA
    Dim prompts As New Dictionary
    With prompts
        .Add "<variable technical name 0>", "<variable value 0>"
        .Add "<variable technical name 1>", "<variable value 1>"
        .Add "<variable technical name 2>", "<variable value 2>"
    End With

    Dim ao As New SapAnalysisOffice
    If Not ao.Refresh("DS_1", prompts) Then
        Debug.Print "Refresh AO fail!"
    End If
```

