---
title: Shapes.AddChart2 Method (Word)
keywords: vbawd10.chm161415273
f1_keywords:
- vbawd10.chm161415273
ms.prod: word
ms.assetid: 54b1e65b-57ad-4824-2acf-2e1e0a22f085
ms.date: 06/08/2017
---


# Shapes.AddChart2 Method (Word)

Adds a chart to the document. Returns a [Shape](shape-object-word.md) object that represents a chart and adds it to the specified collection.


## Syntax

 _expression_ . **AddChart2**_(Style,_ _Type,_ _Left,_ _Top,_ _Width,_ _Height,_ _Anchor,_ _NewLayout)_

 _expression_ A variable that represents a **Shapes** object.


### Parameters



|**Name**|**Required/Optional**|**Data type**|**Description**|
|:-----|:-----|:-----|:-----|
|||||
| _Style_|Optional|INT32|The chart style. Use "-1" to get the default style for the chart type specified in  **Type**.|
| _Type_|Optional|[XLCHARTTYPE](http://msdn.microsoft.com/library/bba4ee89-ee91-f55a-d2e0-59a73e5bfabe%28Office.15%29.aspx)|The type of chart.|
| _Left_|Optional|VARIANT|The position, in points, of the left edge of the chart, relative to the anchor.|
| _Top_|Optional|VARIANT|The position, in points, of the top edge of the chart, relative to the anchor.|
| _Width_|Optional|VARIANT|The width, in points, of the chart.|
| _Height_|Optional|VARIANT|The height, in points, of the chart.|
| _Anchor_|Optional|VARIANT|If  _NewLayout_ is true, the chart will be inserted by using the new dynamic formatting rules (Title is on, and Legend is on only if there are multiple series).|
|||||
|||||
|||||
|||||
|||||
|||||
|||||
| _Type_|Optional|XLCHARTTYPE||
| _Top_|Optional|VARIANT||
| _Width_|Optional|VARIANT||
| _Height_|Optional|VARIANT||
| _Anchor_|Optional|VARIANT||
| _NewLayout_|Optional|VARIANT||

### Return value

 **SHAPE**


## See also


#### Concepts


[Shapes Collection](shapes-object-word.md)

