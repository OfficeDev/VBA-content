---
title: View.ShowAll Property (Word)
keywords: vbawd10.chm161808387
f1_keywords:
- vbawd10.chm161808387
ms.prod: word
api_name:
- Word.View.ShowAll
ms.assetid: 21af8a5b-2110-a2e0-e705-40a66c410625
ms.date: 06/08/2017
---


# View.ShowAll Property (Word)

 **True** if all nonprinting characters (such as hidden text, tab marks, space marks, and paragraph marks) are displayed. Read/write **Boolean** .


## Syntax

 _expression_ . **ShowAll**

 _expression_ Required. A variable that represents a **[View](view-object-word.md)** object.


## Example

This example displays all nonprinting characters in the active window.


```vb
ActiveDocument.ActiveWindow.View.ShowAll = True
```

This example toggles the display of nonprinting characters in the first window.




```vb
Windows(1).View.ShowAll = Not Windows(1).View.ShowAll
```


## See also


#### Concepts


[View Object](view-object-word.md)

