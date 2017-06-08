---
title: Axes.Item Method (Word)
ms.prod: word
api_name:
- Word.Axes.Item
ms.assetid: 143898d3-cbc8-ebfc-4e25-caceeb91a8bf
ms.date: 06/08/2017
---


# Axes.Item Method (Word)

Returns a single  **[Axis](axis-object-word.md)** object from an **Axes** collection.


## Syntax

 _expression_ . **Item**( **_Type_** , **_AxisGroup_** )

 _expression_ A variable that represents an **[Axes](axes-object-word.md)** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Type_|Required| **[XlAxisType](xlaxistype-enumeration-word.md)**|One of the enumeration values that specifies the axis type.|
| _AxisGroup_|Optional| **[XlAxisGroup](xlaxisgroup-enumeration-word.md)**|One of the enumeration values that specifies the axis.|

## Example

The following example sets the title text for the category axis for the first chart in the active document.


```vb
With ActiveDocument.InlineShapes(1) 
 If .HasChart Then 
 With .Chart.Axes.Item(xlCategory) 
 .HasTitle = True 
 .AxisTitle.Caption = "1994" 
 End With 
 End If 
End With
```


## See also


#### Concepts


[Axes Object](axes-object-word.md)

