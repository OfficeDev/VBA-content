---
title: Axis.MinorUnitScale Property (Word)
keywords: vbawd10.chm113049663
f1_keywords:
- vbawd10.chm113049663
ms.prod: word
api_name:
- Word.Axis.MinorUnitScale
ms.assetid: 3ddf49b7-48f2-144f-bf01-3b0c16673b11
ms.date: 06/08/2017
---


# Axis.MinorUnitScale Property (Word)

Returns or sets the minor unit scale value for the category axis when the  **[CategoryType](axis-categorytype-property-word.md)** property is set to **xlTimeScale** . Read/write **[XlTimeUnit](xltimeunit-enumeration-word.md)** .


## Syntax

 _expression_ . **MinorUnitScale**

 _expression_ A variable that represents an **[Axis](axis-object-word.md)** object.


## Remarks

 **MinorUnitScale** can be one of the following **XlTimeUnit** constants:


-  **xlMonths**
    
-  **xlDays**
    
-  **xlYears**
    

## Example

The following example sets the category axis to use a time scale and sets the major and minor units.


```vb
With ActiveDocument.InlineShapes(1) 
 If .HasChart Then 
 With .Chart.Axes(xlCategory) 
 .CategoryType = xlTimeScale 
 .MajorUnit = 5 
 .MajorUnitScale = xlDays 
 .MinorUnit = 1 
 .MinorUnitScale = xlDays 
 End With 
 End If 
End With
```


## See also


#### Concepts


[Axis Object](axis-object-word.md)

