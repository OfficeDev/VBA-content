---
title: View.ShowFieldCodes Property (Word)
keywords: vbawd10.chm161808388
f1_keywords:
- vbawd10.chm161808388
ms.prod: word
api_name:
- Word.View.ShowFieldCodes
ms.assetid: f872636f-9c9f-4dad-d2a0-e18c82d33c68
ms.date: 06/08/2017
---


# View.ShowFieldCodes Property (Word)

 **True** if field codes are displayed. Read/write **Boolean** .


## Syntax

 _expression_ . **ShowFieldCodes**

 _expression_ An expression that returns a **[View](view-object-word.md)** object.


## Example

This example hides field codes in the window for Document1.


```vb
Windows("Document1").View.ShowFieldCodes = False
```

This example shows field codes in the first window.




```vb
Windows(1).View.ShowFieldCodes = True
```

This example toggles field codes in the active window.




```vb
ActiveDocument.ActiveWindow.View.ShowFieldCodes = _ 
 Not ActiveDocument.ActiveWindow.View.ShowFieldCodes
```


## See also


#### Concepts


[View Object](view-object-word.md)

