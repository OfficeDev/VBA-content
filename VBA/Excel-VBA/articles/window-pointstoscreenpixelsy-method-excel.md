---
title: Window.PointsToScreenPixelsY Method (Excel)
keywords: vbaxl10.chm356130
f1_keywords:
- vbaxl10.chm356130
ms.prod: excel
api_name:
- Excel.Window.PointsToScreenPixelsY
ms.assetid: ec25e6d4-22c1-2444-9582-37187901ae02
ms.date: 06/08/2017
---


# Window.PointsToScreenPixelsY Method (Excel)

Converts a vertical measurement from points (document coordinates) to screen pixels (screen coordinates). Returns the converted measurement as a  **Long** value.


## Syntax

 _expression_ . **PointsToScreenPixelsY**( **_Points_** )

 _expression_ A variable that represents a **Window** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Points_|Required| **Long**|The number of points vertically along the left edge of the document window, starting from the top.|

### Return Value

Long


## Example

This example determines the height and width (in pixels) of the selected cells in the active window and returns the values in the  `lWinWidth` and `lWinHeight` variables.


```vb
With ActiveWindow 
 lWinWidth = _ 
 .PointsToScreenPixelsX(.Selection.Width) 
 lWinHeight = _ 
 .PointsToScreenPixelsY(.Selection.Height) 
End With
```


## See also


#### Concepts


[Window Object](window-object-excel.md)

