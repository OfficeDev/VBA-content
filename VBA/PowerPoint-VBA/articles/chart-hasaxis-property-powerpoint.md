---
title: Chart.HasAxis Property (PowerPoint)
keywords: vbapp10.chm684031
f1_keywords:
- vbapp10.chm684031
ms.prod: powerpoint
api_name:
- PowerPoint.Chart.HasAxis
ms.assetid: edb836fb-1a4c-cf70-2ec0-0272b3681e39
ms.date: 06/08/2017
---


# Chart.HasAxis Property (PowerPoint)

Returns or sets which axes exist on the chart. Read/write  **Variant**.


## Syntax

 _expression_. **HasAxis**( **_Index1_**, **_Index2_** )

 _expression_ A variable that represents a **[Chart](chart-object-powerpoint.md)** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Index1_|Optional|**Variant**|The axis type. Series axes apply only to 3-D charts. Can be one of the  **[XlAxisType](xlaxistype-enumeration-powerpoint.md)** constants.|
| _Index2_|Optional|**Variant**|The axis group. 3-D charts have only one set of axes. Can be one of the  **[XlAxisGroup](xlaxisgroup-enumeration-powerpoint.md)** constants.|

## Remarks

You must enter a value for at least one of the parameters when you set this property.

Microsoft Word may create or delete axes if you change the chart type or the  **[AxisGroup](axis-axisgroup-property-powerpoint.md)** property.


## Example




 **Note**  Although the following code applies to Microsoft Word, you can readily modify it to apply to PowerPoint.

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


[Chart Object](chart-object-powerpoint.md)

