---
title: Global.Windows Property (Word)
keywords: vbawd10.chm163119106
f1_keywords:
- vbawd10.chm163119106
ms.prod: word
api_name:
- Word.Global.Windows
ms.assetid: 23ebd91a-8f72-4f63-4ad8-95f98e36309c
ms.date: 06/08/2017
---


# Global.Windows Property (Word)

Returns a  **Windows** collection that represents all open document windows. Read-only.


## Syntax

 _expression_ . **Windows**

 _expression_ A variable that represents a **[Global](global-object-word.md)** object.


## Remarks

For information about returning a single member of a collection, see [Returning an Object from a Collection](http://msdn.microsoft.com/library/28f76384-f495-9640-a7c8-10ada3fac727%28Office.15%29.aspx).


## Example

This example arranges all open windows so that they don't overlap.


```
Windows.Arrange ArrangeStyle:=wdTiled
```


## See also


#### Concepts


[Global Object](global-object-word.md)

