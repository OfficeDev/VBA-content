---
title: Shape.Apply Method (Word)
keywords: vbawd10.chm161480714
f1_keywords:
- vbawd10.chm161480714
ms.prod: word
api_name:
- Word.Shape.Apply
ms.assetid: 3a42c1a6-7037-2649-c079-68f1391521a3
ms.date: 06/08/2017
---


# Shape.Apply Method (Word)

Applies to the specified shape formatting that has been copied using the  **PickUp** method.


## Syntax

 _expression_ . **Apply**

 _expression_ Required. A variable that represents a **[Shape](shape-object-word.md)** object.


## Remarks

If formatting for the  **Shape** object has not previously been copied using the **PickUp** method, using the **Apply** method generates an error.


## Example

This example copies the formatting of shape one on the active document and applies the copied formatting to shape two on the same document.


```vb
With ActiveDocument 
 .Shapes(1).PickUp 
 .Shapes(2).Apply 
End With
```


## See also


#### Concepts


[Shape Object](shape-object-word.md)

