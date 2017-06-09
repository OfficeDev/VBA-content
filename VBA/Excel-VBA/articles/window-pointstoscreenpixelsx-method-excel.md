---
title: Window.PointsToScreenPixelsX Method (Excel)
keywords: vbaxl10.chm356129
f1_keywords:
- vbaxl10.chm356129
ms.prod: excel
api_name:
- Excel.Window.PointsToScreenPixelsX
ms.assetid: b637ae59-30fe-a5cd-2c0d-d9cb63c77d84
ms.date: 06/08/2017
---


# Window.PointsToScreenPixelsX Method (Excel)

Converts a horizontal measurement from points (document coordinates) to screen pixels (screen coordinates). Returns the converted measurement as a  **Long** value.


## Syntax

 _expression_ . **PointsToScreenPixelsX**( **_Points_** )

 _expression_ A variable that represents a **Window** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Points_|Required| **Long**|The number of points horizontally along the top of the document window, starting from the left.|

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

