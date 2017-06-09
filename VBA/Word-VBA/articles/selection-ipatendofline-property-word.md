---
title: Selection.IPAtEndOfLine Property (Word)
keywords: vbawd10.chm158663061
f1_keywords:
- vbawd10.chm158663061
ms.prod: word
api_name:
- Word.Selection.IPAtEndOfLine
ms.assetid: 8db37c0f-6c68-7ccd-0c34-9a40b62b9e9d
ms.date: 06/08/2017
---


# Selection.IPAtEndOfLine Property (Word)

 **True** if the insertion point is at the end of a line that wraps to the next line. Read-only **Boolean** .


## Syntax

 _expression_ . **IPAtEndOfLine**

 _expression_ A variable that represents a **[Selection](selection-object-word.md)** object.


## Remarks

 **False** if the selection isn't collapsed, if the insertion point isn't at the end of a line, or if the insertion point is positioned before a paragraph mark.


## Example

If the insertion point isn't already at the end of the line, this example moves it there.


```vb
Selection.Collapse Direction:=wdCollapseEnd 
If Selection.IPAtEndOfLine = False Then 
 Selection.EndKey Unit:=wdLine, Extend:=wdMove 
End If
```


## See also


#### Concepts


[Selection Object](selection-object-word.md)

