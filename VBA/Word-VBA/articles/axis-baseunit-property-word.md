---
title: Axis.BaseUnit Property (Word)
keywords: vbawd10.chm113049657
f1_keywords:
- vbawd10.chm113049657
ms.prod: word
api_name:
- Word.Axis.BaseUnit
ms.assetid: 1b154779-ac5f-05fc-48d5-cab5ff0f7de7
ms.date: 06/08/2017
---


# Axis.BaseUnit Property (Word)

Returns or sets the base unit for the specified category axis. Read/write  **[XlTimeUnit](xltimeunit-enumeration-word.md)** .


## Syntax

 _expression_ . **BaseUnit**

 _expression_ A variable that represents an **[Axis](axis-object-word.md)** object.


## Remarks

Setting this property has no visible effect if the  **[CategoryType](axis-categorytype-property-word.md)** property for the specified axis is set to **xlCategoryScale** . The set value is retained, however, and takes effect when the **CategoryType** property is set to **xlTimeScale** .

You cannot set this property for a value axis.


## Example

The following example sets the category axis for the first chart in the active document to use a time scale, using months as the base unit.


```vb
 
With ActiveDocument.InlineShapes(1) 
 If .HasChart Then 
 With .Chart 
 .Axes(xlCategory).CategoryType = xlTimeScale 
 .Axes(xlCategory).BaseUnit = xlMonths 
 End With 
 End If 
End With
```


## See also


#### Concepts


[Axis Object](axis-object-word.md)

