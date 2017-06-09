---
title: Application.FeatureInstall Property (Access)
keywords: vbaac10.chm12590
f1_keywords:
- vbaac10.chm12590
ms.prod: access
api_name:
- Access.Application.FeatureInstall
ms.assetid: bc9057bc-72a4-0344-a50a-7b73a2d30212
ms.date: 06/08/2017
---


# Application.FeatureInstall Property (Access)

You can use the  **FeatureInstall** property to specify or determine how Microsoft Access handles calls to methods and properties that require features not yet installed. Read/write **[MsoFeatureInstall](http://msdn.microsoft.com/library/25256738-d169-5c00-1d5d-eb8019811976%28Office.15%29.aspx)**.


## Syntax

 _expression_. **FeatureInstall**

 _expression_ A variable that represents an **Application** object.


## Remarks

When VBA code references an object that is not installed the Microsoft Installer technology will attempt to install the required feature. You use the  **FeatureInstall** property to control what happens when an uninstalled object is referenced. When this property is set to the default, any attempt to use an uninstalled object causes the Installer technology to try to install the requested feature. In some circumstances this may take some time, and the user may believe that the machine has stopped responding to additional commands.

You can set the  **FeatureInstall** property to **msoFeatureInstallOnDemandWithUI** so users can see that something is happening as the feature is being installed. You can set the **FeatureInstall** property to **msoFeatureInstallNone** if you want to trap the error that is returned and display your own dialog box to the user or take some other custom action.

If you have the  **[UserControl](application-usercontrol-property-access.md)** property set to **False**, users will not be prompted to install new features even if the **FeatureInstall** property is set to **msoFeatureInstallOnDemand**. If the **UserControl** property is set to **True**, an installation progress meter will appear if the **FeatureInstall** property is set to **msoFeatureInstallOnDemand**.


## Example

This example checks the value of the  **FeatureInstall** property. If the property is set to **msoFeatureInstallNone**, the code displays a message box that asks the user whether they want to change the property setting. If the user responds "Yes", the property is set to **msoFeatureInstallOnDemand**. The example uses an object variable named MyOfficeApp that is dimensioned as an application object.


```vb
 
 
Dim myofficeapp As Access.Application 
Set myofficeapp = New Access.Application 
 
With MyOfficeApp 
    If .FeatureInstall = msoFeatureInstallNone Then 
        Reply = MsgBox("Uninstalled features for " _ 
            &; "this application may " &; vbCrLf _ 
            &; "cause a run-time error when called." _ 
            &; vbCrLf &; vbCrLf _ 
            &; "Would you like to change this setting" &; vbCrLf _ 
            &; "to automatically install missing features?", _ 
            vbYesNo, "Feature Install Setting") 
            If Reply = vbYes Then 
                .FeatureInstall = msoFeatureInstallOnDemand 
            End If 
    End If 
End With
```


## See also


#### Concepts


[Application Object](application-object-access.md)

