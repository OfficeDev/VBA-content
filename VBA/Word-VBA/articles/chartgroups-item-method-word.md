---
title: ChartGroups.Item Method (Word)
keywords: vbawd10.chm77004800
f1_keywords:
- vbawd10.chm77004800
ms.prod: word
api_name:
- Word.ChartGroups.Item
ms.assetid: 0d78e50d-f2e1-1617-a563-65cc48ca2c30
ms.date: 06/08/2017
---


# ChartGroups.Item Method (Word)

Returns a single object from a collection.


## Syntax

 _expression_ . **Item**( **_Index_** )

 _expression_ A variable that represents a **[ChartGroups](chartgroups-object-word.md)** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Index_|Required| **Variant**|The index number for the object.|

### Return Value

A  **[ChartGroup](chartgroup-object-word.md)** object contained by the collection.


## Example

The following example adds drop lines to chart group one for the first chart group of the first chart in the active document.


```vb
With ActiveDocument.InlineShapes(1) 
 If .HasChart Then 
 .Chart.ChartGroups.Item(1).HasDropLines = True 
 End If 
End With
```


## See also


#### Concepts


[ChartGroups Object](chartgroups-object-word.md)

