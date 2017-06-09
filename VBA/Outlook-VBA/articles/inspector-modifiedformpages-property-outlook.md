---
title: Inspector.ModifiedFormPages Property (Outlook)
keywords: vbaol11.chm2964
f1_keywords:
- vbaol11.chm2964
ms.prod: outlook
api_name:
- Outlook.Inspector.ModifiedFormPages
ms.assetid: ac377d47-846a-1217-592f-7ed190b824ca
ms.date: 06/08/2017
---


# Inspector.ModifiedFormPages Property (Outlook)

Returns the  **[Pages](pages-object-outlook.md)** collection that represents all the pages for the item in the inspector. Read-only.


## Syntax

 _expression_ . **ModifiedFormPages**

 _expression_ A variable that represents an **Inspector** object.


## Remarks

The main page and up to five customizable pages can be obtained using the  **[Add](pages-add-method-outlook.md)** method.


## Example

This Visual Basic for Applications (VBA) displays the count of pages in the  **ModifiedFormPages** collection. To run this example without any errors, display a contact item in the active window.


```vb
Sub CountModifiedFormPages() 
 
 Dim myItem As Outlook.ContactItem 
 
 Dim myPages As Outlook.Pages 
 
 
 
 Set myItem = Application.ActiveInspector.CurrentItem 
 
 Set myPages = myItem.GetInspector.ModifiedFormPages 
 
 MsgBox myPages.Count 
 
End Sub
```


## See also


#### Concepts


[Inspector Object](inspector-object-outlook.md)

