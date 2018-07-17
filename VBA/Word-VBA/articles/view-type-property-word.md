---
title: View.Type Property (Word)
keywords: vbawd10.chm161808384
f1_keywords:
- vbawd10.chm161808384
ms.prod: word
api_name:
- Word.View.Type
ms.assetid: 0168c7cd-147f-b81b-2a56-3c3f751cc4b0
ms.date: 06/08/2017
---


# View.Type Property (Word)

Returns or sets the view type. Read/write  **[WdViewType](wdviewtype-enumeration-word.md)** .


## Syntax

 _expression_ . **Type**

 _expression_ Required. A variable that represents a **[View](view-object-word.md)** object.


## Remarks

The  **Type** property returns **wdMasterView** for all documents where the current view is an outline or a master document. The current view will never return **wdOutlineView** unless explicitly set first in code.

To check whether the current document is an outline, use the  **Type** property and the **Subdocuments** collection's **Count** property. If the **Type** property returns either **wdOutlineView** or **wdMasterView** and the **Count** property returns zero, the document is an outline. For example:






```vb
Sub VerifyOutlineView() 
 With ActiveWindow.View 
 If .Type = wdOutlineView Or wdMasterView Then 
 If ActiveDocument.Subdocuments.Count = 0 Then 
 . 
 . 
 . 
 End If 
 End If 
 End With 
End Sub
```


## Example

This example switches the active window to print preview. The  **Type** property creates a new print preview window.


```vb
ActiveDocument.ActiveWindow.View.Type = wdPrintPreview
```


## See also


#### Concepts


[View Object](view-object-word.md)

