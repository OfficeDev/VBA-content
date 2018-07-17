---
title: View.ShowFirstLineOnly Property (Word)
keywords: vbawd10.chm161808392
f1_keywords:
- vbawd10.chm161808392
ms.prod: word
api_name:
- Word.View.ShowFirstLineOnly
ms.assetid: 16b67deb-e65d-10ac-f856-4f7df0a4ccbc
ms.date: 06/08/2017
---


# View.ShowFirstLineOnly Property (Word)

 **True** if only the first line of body text is shown in outline view. Read/write **Boolean** .


## Syntax

 _expression_ . **ShowFirstLineOnly**

 _expression_ An expression that returns a **[View](view-object-word.md)** object.


## Remarks

This property generates an error if the view isn't outline or master document view.


## Example

This example switches the active window to outline view and hides all but the first line of body text.


```vb
With ActiveDocument.ActiveWindow.View 
 .Type = wdOutlineView 
 .ShowFirstLineOnly = True 
End With
```


## See also


#### Concepts


[View Object](view-object-word.md)

