---
title: Inspector.SetCurrentFormPage Method (Outlook)
keywords: vbaol11.chm2969
f1_keywords:
- vbaol11.chm2969
ms.prod: outlook
api_name:
- Outlook.Inspector.SetCurrentFormPage
ms.assetid: a0e11ca9-d5be-cec9-ad78-bfbaec1b92d6
ms.date: 06/08/2017
---


# Inspector.SetCurrentFormPage Method (Outlook)

Displays the specified form page or form region in the inspector.


## Syntax

 _expression_ . **SetCurrentFormPage**( **_PageName_** )

 _expression_ A variable that represents an **Inspector** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _PageName_|Required| **String**|The display name of the form page, or the internal name of a form region.|

## Remarks

You can use  **SetCurrentFormPage** to display a form region by specifying the **[InternalName](formregion-internalname-property-outlook.md)** property of the form region, if the form region is an a separate, replace, or replace-all form region.


## Example

This Visual Basic for Applications (VBA) example uses the  **SetCurrentFormPage** method to show the **All Fields** page of the currently open item. If an error occurs, Outlook will display a message box to the user.


```vb
Sub ShowAllFieldsPage() 
 
 On Error GoTo ErrorHandler 
 
 Dim myInspector As Inspector 
 
 Dim myItem As Object 
 
 
 
 Set myInspector = Application.ActiveInspector 
 
 myInspector.SetCurrentFormPage ("All Fields") 
 
 Set myItem = myInspector.CurrentItem 
 
 myItem.Display 
 
Exit Sub 
 
 
 
ErrorHandler: 
 
 MsgBox Err.Description, vbInformation 
 
End Sub
```


## See also


#### Concepts


[Inspector Object](inspector-object-outlook.md)

