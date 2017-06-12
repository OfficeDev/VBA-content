---
title: Chart.SetDefaultChart Method (Word)
keywords: vbawd10.chm79364171
f1_keywords:
- vbawd10.chm79364171
ms.prod: word
api_name:
- Word.Chart.SetDefaultChart
ms.assetid: e914b44a-5de9-ca9d-a513-96943802a194
ms.date: 06/08/2017
---


# Chart.SetDefaultChart Method (Word)

Specifies the name of the chart template that Microsoft Word uses when it creates new charts.


## Syntax

 _expression_ . **SetDefaultChart**( **_Name_** )

 _expression_ A variable that represents a **[Chart](chart-object-word.md)** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Name_|Required| **Variant**|Specifies the name of the default chart template that Word uses when it creates new charts. This name can be set to either the name of a user-defined chart template in the gallery or a special  **[XlChartGallery](xlchartgallery-enumeration-word.md)** constant, **xlBuiltIn** , to specify a built-in chart template.|

## Example

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


[Chart Object](chart-object-word.md)

