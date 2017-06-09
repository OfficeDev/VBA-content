---
title: View.FullScreen Property (Word)
keywords: vbawd10.chm161808385
f1_keywords:
- vbawd10.chm161808385
ms.prod: word
api_name:
- Word.View.FullScreen
ms.assetid: f26be86f-be78-84f7-8222-187040d61a40
ms.date: 06/08/2017
---


# View.FullScreen Property (Word)

 **True** if the window is in full-screen view. Read/write **Boolean** .


## Syntax

 _expression_ . **FullScreen**

 _expression_ A variable that represents a **[View](view-object-word.md)** object.


## Example

This example switches the active window to full-screen view.


```vb
ActiveDocument.ActiveWindow.View.FullScreen = True
```

This example activates the window for Sales.doc and switches out of full-screen view.




```vb
With Windows("Sales.doc") 
 .Activate 
 .View.FullScreen = False 
End With
```


## See also


#### Concepts


[View Object](view-object-word.md)

