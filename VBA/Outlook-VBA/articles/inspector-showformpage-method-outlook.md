---
title: Inspector.ShowFormPage Method (Outlook)
keywords: vbaol11.chm2970
f1_keywords:
- vbaol11.chm2970
ms.prod: outlook
api_name:
- Outlook.Inspector.ShowFormPage
ms.assetid: d31a4df6-7b94-5eb4-8ec9-5a03dcaae53a
ms.date: 06/08/2017
---


# Inspector.ShowFormPage Method (Outlook)

Displays a button in the  **Show** group of the Microsoft Office Fluent ribbon for the inspector, clicking which shows the page or form region specified by _PageName_.


## Syntax

 _expression_ . **ShowFormPage**( **_PageName_** )

 _expression_ A variable that represents an **Inspector** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _PageName_|Required| **String**|The display name of the form page, or the internal name of a form region to be shown.|

## Remarks

For form regions, you can use  **ShowFormRegion** to display the button by specifying the **[InternalName](formregion-internalname-property-outlook.md)** property of the form region, if the form region is an adjoining or separate form region. Only the add-in that implements the form region can use **ShowFormRegion** to display the button.


## Example

This Visual Basic for Applications (VBA) example uses the  **ShowFormPage** method to show a button, labeled **All Fields**, in the  **Show** group of the ribbon of the active inspector. Clicking the **All Fields** button displays the **All Fields** page of the currently open item. If an error occurs, Outlook will display a message box to the user.


```vb
Sub ShowAllFieldsPage() 
 
 On Error GoTo ErrorHandler 
 
 Dim myInspector As Outlook.Inspector 
 
 Dim myItem As Object 
 
 
 
 Set myInspector = Application.ActiveInspector 
 
 myInspector.ShowFormPage ("All Fields") 
 
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

