---
title: Chart.HasAxis Property (Word)
keywords: vbawd10.chm79364150
f1_keywords:
- vbawd10.chm79364150
ms.prod: word
api_name:
- Word.Chart.HasAxis
ms.assetid: b5b7effe-48c6-75d9-fdc4-7a9ff148f0e9
ms.date: 06/08/2017
---


# Chart.HasAxis Property (Word)

Returns or sets which axes exist on the chart. Read/write  **Variant** .


## Syntax

 _expression_ . **HasAxis**( **_Index1_** , **_Index2_** )

 _expression_ A variable that represents a **[Chart](chart-object-word.md)** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Index1_|Optional| **Variant**|The axis type. Series axes apply only to 3-D charts. Can be one of the  **[XlAxisType](xlaxistype-enumeration-word.md)** constants.|
| _Index2_|Optional| **Variant**|The axis group. 3-D charts have only one set of axes. Can be one of the  **[XlAxisGroup](xlaxisgroup-enumeration-word.md)** constants.|

## Remarks

You must enter a value for at least one of the parameters when you set this property.

Microsoft Word may create or delete axes if you change the chart type or the  **[AxisGroup](axis-axisgroup-property-word.md)** property.


## Example

The following example enables the primary value axis for the first chart in the active document.


```vb
With ActiveDocument.InlineShapes(1) 
 If .HasChart Then 
 .Chart.HasAxis(xlValue, xlPrimary) = True 
 End If 
End With 

```


## See also


#### Concepts


[Chart Object](chart-object-word.md)

