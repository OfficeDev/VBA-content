---
title: View.ShowFormat Property (Word)
keywords: vbawd10.chm161808393
f1_keywords:
- vbawd10.chm161808393
ms.prod: word
api_name:
- Word.View.ShowFormat
ms.assetid: 8171ff9b-5e5d-a3c1-2ea0-31743991ea8e
ms.date: 06/08/2017
---


# View.ShowFormat Property (Word)

 **True** if character formatting is visible in outline view. Read/write **Boolean** .


## Syntax

 _expression_ . **ShowFormat**

 _expression_ An expression that returns a **[View](view-object-word.md)** object.


## Remarks

This property generates an error if the view isn't outline or master document view.


## Example

This example switches the active window to outline view and shows character formatting.


```vb
With ActiveDocument.ActiveWindow.View 
 .Type = wdOutlineView 
 .ShowFormat = True 
End With
```


## See also


#### Concepts


[View Object](view-object-word.md)

