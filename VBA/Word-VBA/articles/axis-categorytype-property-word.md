---
title: Axis.CategoryType Property (Word)
keywords: vbawd10.chm113049665
f1_keywords:
- vbawd10.chm113049665
ms.prod: word
api_name:
- Word.Axis.CategoryType
ms.assetid: 891a0cce-f5cb-6a8a-6216-fa6aaa1adac9
ms.date: 06/08/2017
---


# Axis.CategoryType Property (Word)

Returns or sets the category axis type. Read/write  **[XlCategoryType](xlcategorytype-enumeration-word.md)** .


## Syntax

 _expression_ . **CategoryType**

 _expression_ A variable that represents an **[Axis](axis-object-word.md)** object.


## Remarks

You cannot set this property for a value axis.


## Example

The following example sets the category axis for the first chart in the active document to use a time scale, using months as the base unit.


```vb
With ActiveDocument.InlineShapes(1) 
 If .HasChart Then 
 With .Chart.Axes(xlCategory) 
 .CategoryType = xlTimeScale 
 .BaseUnit = xlMonths 
 End With 
 End If 
End With
```


## See also


#### Concepts


[Axis Object](axis-object-word.md)

