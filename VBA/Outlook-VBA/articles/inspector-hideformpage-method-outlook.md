---
title: Inspector.HideFormPage Method (Outlook)
keywords: vbaol11.chm2967
f1_keywords:
- vbaol11.chm2967
ms.prod: outlook
api_name:
- Outlook.Inspector.HideFormPage
ms.assetid: fbb0fec9-5a23-50f8-0be6-3d264859f327
ms.date: 06/08/2017
---


# Inspector.HideFormPage Method (Outlook)

Hides a form page or a form region in the inspector.


## Syntax

 _expression_ . **HideFormPage**( **_PageName_** )

 _expression_ A variable that represents an **Inspector** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _PageName_|Required| **String**|The display name of the form page, or the internal name of a form region to be hidden.|

## Remarks

You can use  **HideFormRegion** to hide a form region by specifying the **[InternalName](formregion-internalname-property-outlook.md)** property of the form region, if the form region is an adjoining or separate form region. Only the add-in that implements the form region can hide the form region.


## Example

This Visual Basic for Applications (VBA) example uses  **HideFormPage** to hide the "General" page of a newly-created **[ContactItem](contactitem-object-outlook.md)** and displays the item.


```vb
Sub HidePage() 
 
 Dim MyItem As Outlook.ContactItem 
 
 Dim myPages As Outlook.Pages 
 
 Dim myinspector As Outlook.Inspector 
 
 
 
 Set MyItem = Application.CreateItem(olContactItem) 
 
 Set myPages = MyItem.GetInspector.ModifiedFormPages 
 
 myPages.Add "General" 
 
 Set myinspector = Application.ActiveInspector 
 
 myinspector.HideFormPage "General" 
 
 MyItem.Display 
 
End Sub
```


## See also


#### Concepts


[Inspector Object](inspector-object-outlook.md)

