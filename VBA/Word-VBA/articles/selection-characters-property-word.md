---
title: Selection.Characters Property (Word)
keywords: vbawd10.chm158662709
f1_keywords:
- vbawd10.chm158662709
ms.prod: word
api_name:
- Word.Selection.Characters
ms.assetid: 605c0fc5-f5b9-6782-9fdd-54589040d243
ms.date: 06/08/2017
---


# Selection.Characters Property (Word)

Returns a  **[Characters](characters-object-word.md)** collection that represents the characters in a document, range, or selection. Read-only.


## Syntax

 _expression_ . **Characters**

 _expression_ A variable that represents a **[Selection](selection-object-word.md)** object.


## Remarks

For information about returning a single member of a collection, see [Returning an Object from a Collection](http://msdn.microsoft.com/library/28f76384-f495-9640-a7c8-10ada3fac727%28Office.15%29.aspx).


## Example

This example displays the first character in the selection. If nothing is selected, the character immediately after the insertion point is displayed.


```
char = Selection.Characters(1).Text 
MsgBox "The first character is... " &; char
```


## See also


#### Concepts


[Selection Object](selection-object-word.md)

