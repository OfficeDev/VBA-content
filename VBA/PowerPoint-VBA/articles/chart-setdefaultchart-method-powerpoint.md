---
title: Chart.SetDefaultChart Method (PowerPoint)
keywords: vbapp10.chm684006
f1_keywords:
- vbapp10.chm684006
ms.prod: powerpoint
api_name:
- PowerPoint.Chart.SetDefaultChart
ms.assetid: a75ac074-dd5d-7530-2446-cc89b3d1ac5f
ms.date: 06/08/2017
---


# Chart.SetDefaultChart Method (PowerPoint)

Specifies the name of the chart template that Microsoft Word uses when it creates new charts.


## Syntax

 _expression_. **SetDefaultChart**( **_Name_** )

 _expression_ A variable that represents a **[Chart](chart-object-powerpoint.md)** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Name_|Required|**Variant**|Specifies the name of the default chart template that Word uses when it creates new charts. This name can be set to either the name of a user-defined chart template in the gallery or a special  **[XlChartGallery](xlchartgallery-enumeration-powerpoint.md)** constant, **xlBuiltIn**, to specify a built-in chart template.|

## Example




 **Note**  Although the following code applies to Microsoft Word, you can readily modify it to apply to PowerPoint.

The following example sets the default chart template to a custom chart template named "Monthly Sales."




```vb
With ActiveDocument.InlineShapes(1)

    If .HasChart Then

        .Chart.SetDefaultChart Name:="Monthly Sales"

    End If

End With
```


## See also


#### Concepts


[Chart Object](chart-object-powerpoint.md)

