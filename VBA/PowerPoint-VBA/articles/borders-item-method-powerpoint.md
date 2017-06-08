---
title: Borders.Item Method (PowerPoint)
keywords: vbapp10.chm629003
f1_keywords:
- vbapp10.chm629003
ms.prod: powerpoint
api_name:
- PowerPoint.Borders.Item
ms.assetid: fad023e2-55c1-4115-fc61-cd4519486fad
ms.date: 06/08/2017
---


# Borders.Item Method (PowerPoint)

Returns a  **[LineFormat](lineformat-object-powerpoint.md)** object for the specified border from the **Borders** collection.


## Syntax

 _expression_. **Item**( **_BorderType_** )

 _expression_ A variable that represents a **Borders** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _BorderType_|Required|**PpBorderType**|Specifies which border of a cell or cell range is to be returned.|

### Return Value

LineFormat


## Remarks

The  _BorderType_ parameter value can be one of these **PpBorderType** constants.


||
|:-----|
|**ppBorderBottom**|
|**ppBorderDiagonalDown**|
|**ppBorderDiagonalUp**|
|**ppBorderLeft**|
|**ppBorderRight**|
|**ppBorderTop**|

## See also


#### Concepts


[Borders Object](borders-object-powerpoint.md)

