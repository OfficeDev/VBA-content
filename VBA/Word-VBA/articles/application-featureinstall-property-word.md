---
title: Application.FeatureInstall Property (Word)
keywords: vbawd10.chm158335423
f1_keywords:
- vbawd10.chm158335423
ms.prod: word
api_name:
- Word.Application.FeatureInstall
ms.assetid: 4abb8363-dee0-0209-2e56-54cea87fe692
ms.date: 06/08/2017
---


# Application.FeatureInstall Property (Word)

Returns or sets how Microsoft Word handles calls to methods and properties that require features not yet installed. Read/write  **MsoFeatureInstall** .


## Syntax

 _expression_ . **FeatureInstall**

 _expression_ Required. A variable that represents an **[Application](application-object-word.md)** object.


## Remarks

You can use the  **msoFeatureInstallOnDemandWithUI** constant to prevent users from believing that the application isn't responding while a feature is being installed. Use the **msoFeatureInstallNone** constant if you want the developer to be the only one who can install features.

If you have the  **DisplayAlerts** property set to **False** , users will not be prompted to install new features even if the **FeatureInstall** property is set to **msoFeatureInstallOnDemand** . If the **DisplayAlerts** property is set to **True** , an installation progress meter will appear if the **FeatureInstall** property is set to **msoFeatureInstallOnDemand** .


## Example

This example activates a new instance of Microsoft Excel and checks the value of the  **FeatureInstall** property. If the property is set to **msoFeatureInstallNone** , the code displays a message box that asks the user whether they want to change the property setting. If the user responds "Yes," the property is set to **msoFeatureInstallOnDemand** .


 **Note**  For this example to function properly, you must add a reference to Microsoft Excel Object Library.


```vb
Dim ExcelApp As New Excel.Application 
Dim intReply As Integer 
 
With ExcelApp 
 If .FeatureInstall = msoFeatureInstallNone Then 
 intReply = MsgBox("Uninstalled features for " _ 
 &; "this application may " &; vbCrLf _ 
 &; "cause a run-time error when called." _ 
 &; vbCrLf &; vbCrLf _ 
 &; "Would you like to change this setting" &; vbCrLf _ 
 &; "to automatically install missing features?", _ 
 vbYesNo, "Feature Install Setting") 
 If intReply = vbYes Then 
 .FeatureInstall = msoFeatureInstallOnDemand 
 End If 
 End If 
End With
```


## See also


#### Concepts


[Application Object](application-object-word.md)

