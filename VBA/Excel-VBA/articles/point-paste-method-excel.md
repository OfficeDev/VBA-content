---
title: Point.Paste Method (Excel)
keywords: vbaxl10.chm576090
f1_keywords:
- vbaxl10.chm576090
ms.prod: excel
api_name:
- Excel.Point.Paste
ms.assetid: 0a984f1c-54de-d49f-8677-43d513a0f9fc
ms.date: 06/08/2017
---


# Point.Paste Method (Excel)

Pastes a picture from the Clipboard as the marker on the selected point.


## Syntax

 _expression_ . **Paste**

 _expression_ A variable that represents a **Point** object.


### Return Value

Variant


## Remarks

This method can be used on column, bar, line, or radar charts, and it sets the  **[MarkerStyle](point-markerstyle-property-excel.md)** property to **xlMarkerStylePicture** .


## Example

This example pastes a picture from the Clipboard into series one in Chart1.


```vb
Charts("Chart1").SeriesCollection(1).Paste
```


## See also


#### Concepts


[Point Object](point-object-excel.md)

