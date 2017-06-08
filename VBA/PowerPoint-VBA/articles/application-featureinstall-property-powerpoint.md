---
title: Application.FeatureInstall Property (PowerPoint)
keywords: vbapp10.chm502043
f1_keywords:
- vbapp10.chm502043
ms.prod: powerpoint
api_name:
- PowerPoint.Application.FeatureInstall
ms.assetid: 254fc432-9ee5-d978-19ac-5fa6f94daa94
ms.date: 06/08/2017
---


# Application.FeatureInstall Property (PowerPoint)

Returns or sets how Microsoft PowerPoint handles calls to methods and properties that require features not yet installed. Read/write.


## Syntax

 _expression_. **FeatureInstall**

 _expression_ A variable that represents an **Application** object.


### Return Value

MsoFeatureInstall


## Remarks

You can use the  **msoFeatureInstallOnDemandWithUI** constant to prevent users from believing that the application is not responding while a feature is being installed. Use the **msoFeatureInstallNone** constant with error trapping routines to exclude end-user feature installation.


 **Note**  If you refer to an uninstalled presentation design template in a string, a run-time error is generated. The template is not installed automatically regardless of your  **FeatureInstall** property setting. To use the **[ApplyTemplate](presentation-applytemplate-method-powerpoint.md)** method for a template that is not currently installed, you first must install the additional design templates. To do so, install the Additional Design Templates for PowerPoint by running the Microsoft Office installation program (available by clicking the **Add/Remove Programs** icon in Windows Control Panel).

The value of the  **FeatureInstall** property can be one of these **MsoFeatureInstall** constants.



|**Constant**|**Description**|
|:-----|:-----|
|**msoFeatureInstallNone**| The default. A trappable run-time automation error is generated when uninstalled features are called.|
|**msoFeatureInstallOnDemand**| A dialog box is displayed prompting the user to install new features.|
|**msoFeatureInstallOnDemandWithUI**| A progress meter is displayed during installation. The user is not prompted to install new features.|

## Example

This example checks the value of the  **FeatureInstall** property. If the property is set to **msoFeatureInstallNone**, the code displays a message box that asks the user whether they want to change the property setting. If the user responds "Yes", the property is set to **msoFeatureInstallOnDemand**.


```vb
With Application
    If .FeatureInstall = msoFeatureInstallNone Then
        Reply = MsgBox("Uninstalled features for this " _
                &; "application " &; vbCrLf _
                &; "may cause a run-time error when called." &; vbCrLf _
                &; vbCrLf _
                &; "Would you like to change this setting" &; vbCrLf _
                &; "to automatically install missing features when called?" _
                , 52, "Feature Install Setting")

            If Reply = 6 Then
                .FeatureInstall = msoFeatureInstallOnDemand
            End If
    End If
End With
```


## See also


#### Concepts


[Application Object](application-object-powerpoint.md)

