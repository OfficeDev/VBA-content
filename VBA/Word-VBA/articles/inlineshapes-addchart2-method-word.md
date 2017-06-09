---
title: InlineShapes.AddChart2 Method (Word)
keywords: vbawd10.chm162070638
f1_keywords:
- vbawd10.chm162070638
ms.prod: word
ms.assetid: 108899b6-24bb-cf4c-db95-066219536c19
ms.date: 06/08/2017
---


# InlineShapes.AddChart2 Method (Word)

Adds a chart to the document. Returns an [InlineShape](inlineshape-object-word.md) object that represents the chart and adds it to the specified collection.


## Syntax

 _expression_ . **AddChart2**_(Style,_ _Type,_ _Range,_ _NewLayout)_

 _expression_ A variable that represents a **InlineShapes** object.


### Parameters



|**Name**|**Required/Optional**|**Data type**|**Description**|
|:-----|:-----|:-----|:-----|
|||||
| _Style_|Optional|INT32|The chart style. Use "-1" to get the default style for the chart type specified in  **Type**.|
| _Type_|Optional|[XLCHARTTYPE](http://msdn.microsoft.com/library/bba4ee89-ee91-f55a-d2e0-59a73e5bfabe%28Office.15%29.aspx)|The type of chart.|
| _Range_|Optional|VARIANT|The range where the chart will be placed in the text. The chart replaces the range, unless the range is collapsed. If this argument is omitted, the chart is placed automatically.|
| _NewLayout_|Optional|VARIANT|If  _NewLayout_ is true, the chart is inserted by using the new dynamic formatting rules (Title is on, and Legend is on only if there are multiple series).|
| _Type_|Optional|XLCHARTTYPE||

### Return value

 **INLINESHAPE**


## See also


#### Concepts


[InlineShapes Collection](inlineshapes-object-word.md)

