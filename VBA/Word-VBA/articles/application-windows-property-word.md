---
title: Application.Windows Property (Word)
keywords: vbawd10.chm158334978
f1_keywords:
- vbawd10.chm158334978
ms.prod: word
api_name:
- Word.Application.Windows
ms.assetid: 860d9e12-4c02-be1f-64a7-ef0305881854
ms.date: 06/08/2017
---


# Application.Windows Property (Word)

Returns a  **[Windows](windows-object-word.md)** collection that represents all document windows. Read-only.


## Syntax

 _expression_ . **Windows**

 _expression_ A variable that represents an **[Application](application-object-word.md)** object.


## Remarks

The collection corresponds to the window names that appear at the bottom of the Window menu. For information about returning a single member of a collection, see [Returning an Object from a Collection](http://msdn.microsoft.com/library/28f76384-f495-9640-a7c8-10ada3fac727%28Office.15%29.aspx).


## Example

This example arranges all open windows so that they don't overlap.


```
Windows.Arrange ArrangeStyle:=wdTiled
```


## See also


#### Concepts


[Application Object](application-object-word.md)

