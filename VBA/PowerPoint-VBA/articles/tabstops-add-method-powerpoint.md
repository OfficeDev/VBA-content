---
title: TabStops.Add Method (PowerPoint)
keywords: vbapp10.chm573005
f1_keywords:
- vbapp10.chm573005
ms.prod: powerpoint
api_name:
- PowerPoint.TabStops.Add
ms.assetid: cbb8f77f-c5c2-4573-abbe-ddca9bdbdf13
ms.date: 06/08/2017
---


# TabStops.Add Method (PowerPoint)

Creates a tab stop and adds it to the  **TabStops** collection.


## Syntax

 _expression_. **Add**( **_Type_**, **_Position_** )

 _expression_ A variable that represents a **TabStops** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Type_|Required|**PpTabStopType**|The type of the tab stop to be added.|
| _Position_|Required|**Single**|The position of the tab stop in the tab stops collection.|

### Return Value

TabStop


## Remarks

The  _Type_ parameter value can be one of these **PpTabStopType** constants.



|**Constant**|**Description**|
|:-----|:-----|
|**ppTabStopCenter**|Center tab stop.|
|**ppTabStopDecimal**|Decimal tab stop.|
|**ppTabStopLeft**|Left tab stop.|
|**ppTabStopMixed**|Mixed tab stop.|
|**ppTabStopRight**|Right tab stop.|

## See also


#### Concepts


[TabStops Object](tabstops-object-powerpoint.md)

