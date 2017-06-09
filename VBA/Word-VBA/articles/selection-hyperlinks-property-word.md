---
title: Selection.Hyperlinks Property (Word)
keywords: vbawd10.chm158662812
f1_keywords:
- vbawd10.chm158662812
ms.prod: word
api_name:
- Word.Selection.Hyperlinks
ms.assetid: c90c3779-cbb9-4174-3002-850750b4bb41
ms.date: 06/08/2017
---


# Selection.Hyperlinks Property (Word)

Returns a  **[Hyperlinks](hyperlinks-object-word.md)** collection that represents all the hyperlinks in the specified selection. Read-only.


## Syntax

 _expression_ . **Hyperlinks**

 _expression_ A variable that represents a **[Selection](selection-object-word.md)** object.


## Remarks

For information about returning a single member of a collection, see [Returning an Object from a Collection](http://msdn.microsoft.com/library/28f76384-f495-9640-a7c8-10ada3fac727%28Office.15%29.aspx).


## Example

This example jumps to the address of the first hyperlink in the selection.


```vb
If Selection.Hyperlinks.Count >= 1 Then 
 Selection.Hyperlinks(1).Follow 
End If
```


## See also


#### Concepts


[Selection Object](selection-object-word.md)

