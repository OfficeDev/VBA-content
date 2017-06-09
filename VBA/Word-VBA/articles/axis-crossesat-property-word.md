---
title: Axis.CrossesAt Property (Word)
keywords: vbawd10.chm113049608
f1_keywords:
- vbawd10.chm113049608
ms.prod: word
api_name:
- Word.Axis.CrossesAt
ms.assetid: 720fd3a6-89fb-bb55-9b0b-d6ecb2e5ca21
ms.date: 06/08/2017
---


# Axis.CrossesAt Property (Word)

Returns or sets the point on the value axis where the category axis crosses it. Applies only to the value axis. Read/write  **Double** .


## Syntax

 _expression_ . **CrossesAt**

 _expression_ A variable that represents an **[Axis](axis-object-word.md)** object.


## Remarks

Setting this property causes the  **[Crosses](axis-crosses-property-word.md)** property to change to **xlAxisCrossesCustom** . **xlAxisCrossesCustom** is a constant in the **XlAxisCrosses** enumeration.

You cannot use this property on radar charts. For 3-D charts, this property indicates where the plane defined by the category axes crosses the value axis.


## See also


#### Concepts


[Axis Object](axis-object-word.md)

