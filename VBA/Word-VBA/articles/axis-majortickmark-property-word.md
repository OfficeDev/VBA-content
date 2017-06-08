---
title: Axis.MajorTickMark Property (Word)
keywords: vbawd10.chm113049618
f1_keywords:
- vbawd10.chm113049618
ms.prod: word
api_name:
- Word.Axis.MajorTickMark
ms.assetid: f2e4c509-0736-44bd-249b-1963ac697ee4
ms.date: 06/08/2017
---


# Axis.MajorTickMark Property (Word)

Returns or sets the type of major tick mark for the specified axis. Read/write  **[XlTickMark](xltickmark-enumeration-word.md)** .


## Syntax

 _expression_ . **MajorTickMark**

 _expression_ A variable that represents an **[Axis](axis-object-word.md)** object.


## Remarks

 **MajorTickMark** can be set to one of the following **XlTickMark** constants:


-  **xlTickMarkInside**
    
-  **xlTickMarkOutside**
    
-  **xlTickMarkCross**
    
-  **xlTickMarkNone**
    

## Example

The following example sets the major tick marks for the value axis for the first chart in the active document to be outside the axis.


```vb
With ActiveDocument.InlineShapes(1) 
 If .HasChart Then 
 .Chart.Axes(xlValue).MajorTickMark = xlTickMarkOutside 
 End If 
End With 

```


## See also


#### Concepts


[Axis Object](axis-object-word.md)

