---
title: Axis.MaximumScaleIsAuto Property (Word)
keywords: vbawd10.chm113049630
f1_keywords:
- vbawd10.chm113049630
ms.prod: word
api_name:
- Word.Axis.MaximumScaleIsAuto
ms.assetid: 7ec9d4da-0851-146c-2324-bcaba7434158
ms.date: 06/08/2017
---


# Axis.MaximumScaleIsAuto Property (Word)

 **True** if Microsoft Word calculates the maximum value for the value axis. Read/write **Boolean** .


## Syntax

 _expression_ . **MaximumScaleIsAuto**

 _expression_ A variable that represents an **[Axis](axis-object-word.md)** object.


## Remarks

Setting the  **[MaximumScale](axis-maximumscale-property-word.md)** property sets this property to **False** .


## Example

The following example automatically calculates the minimum scale and the maximum scale for the value axis of the first chart in the active document.


```vb
With ActiveDocument.InlineShapes(1) 
 If .HasChart Then 
 With .Chart.Axes(xlValue) 
 .MinimumScaleIsAuto = True 
 .MaximumScaleIsAuto = True 
 End With 
 End If 
End With 

```


## See also


#### Concepts


[Axis Object](axis-object-word.md)

