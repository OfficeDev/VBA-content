---
title: Application.FeatureInstall Property (Excel)
keywords: vbaxl10.chm133259
f1_keywords:
- vbaxl10.chm133259
ms.prod: excel
api_name:
- Excel.Application.FeatureInstall
ms.assetid: 0bfe9d01-543c-9ea8-8ff6-2032f056b070
ms.date: 06/08/2017
---


# Application.FeatureInstall Property (Excel)

Returns or sets a value (constant) that specifies how Microsoft Excel handles calls to methods and properties that require features that aren't yet installed. Can be one of the  **[MsoFeatureInstall](http://msdn.microsoft.com/library/25256738-d169-5c00-1d5d-eb8019811976%28Office.15%29.aspx)** constants listed in the following table. Read/write **MsoFeatureInstall** .


## Syntax

 _expression_ . **FeatureInstall**

 _expression_ A variable that represents an **Application** object.


## Remarks



| **MsoFeatureInstall** can be one of these **MsoFeatureInstall** constants.|
| **msoFeatureInstallNone** . Generates a generic Automation error at run time when uninstalled features are called. This is the default constant.|
| **msoFeatureInstallOnDemand** . Prompts the user to install new features.|
| **msoFeatureInstallOnDemandWithUI** . Displays a progress meter during installation; doesn't prompt the user to install new features.|
You can use the  **msoFeatureInstallOnDemandWithUI** constant to prevent users from thinking that the application isn't responding while a feature is being installed. Use the **msoFeatureInstallNone** constant if you want the developer to be the only one who can install features.

If you have the  **[DisplayAlerts](application-displayalerts-property-excel.md)** property set to **False** , users won't be prompted to install new features even if the **FeatureInstall** property is set to **msoFeatureInstallOnDemand** . If the **DisplayAlerts** property is set to **True** , an installation progress meter will appear if the **FeatureInstall** property is set to **msoFeatureInstallOnDemand** .


## Example

This example activates a new instance of Microsoft Word and checks the value of the  **FeatureInstall** property. Be sure to set a reference to the Microsoft Word object library. If the **FeatureInstall** property is set to **msoFeatureInstallNone** , the code displays a message box that asks the user whether they want to change the property setting. If the user responds Yes, the property is set to **msoFeatureInstallOnDemand** .


```vb
Dim WordApp As New Word.Application, Reply As Integer 
Application.ActivateMicrosoftApp xlMicrosoftWord With WordApp 
    If .FeatureInstall = msoFeatureInstallNone Then 
        Reply = MsgBox("Uninstalled features for this " _ 
            &; "application " &; vbCrLf _ 
            &; "may cause a run-time error when called." &; vbCrLf _ 
            &; vbCrLf _ 
            &; "Would you like to change this setting" &; vbCrLf _ 
            &; "to automatically install missing features?" _ 
            , 52, "Feature Install Setting") 
        If Reply = 6 Then 
            .FeatureInstall = msoFeatureInstallOnDemand 
        End If 
    End If 
End With
```


## See also


#### Concepts


[Application Object](application-object-excel.md)

