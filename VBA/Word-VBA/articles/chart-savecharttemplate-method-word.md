---
title: Chart.SaveChartTemplate Method (Word)
keywords: vbawd10.chm79364173
f1_keywords:
- vbawd10.chm79364173
ms.prod: word
api_name:
- Word.Chart.SaveChartTemplate
ms.assetid: d980f663-7e73-7b55-9f7c-1fc9da84c0bd
ms.date: 06/08/2017
---


# Chart.SaveChartTemplate Method (Word)

Saves a custom chart template to the list of available chart templates.


## Syntax

 _expression_ . **SaveChartTemplate**( **_FileName_** )

 _expression_ A variable that represents a **[Chart](chart-object-word.md)** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _FileName_|Required| **String**|The name of the chart template.|

## Remarks

By default, this method saves the active chart to the user's chart template directory. If a UNC or URL is specified, the chart will be saved to the specified location instead. 


## Example

The following example adds a new chart template based on the first chart of the active document.


```vb
With ActiveDocument.InlineShapes(1) 
 If .HasChart Then 
 .Chart.SaveChartTemplate _ 
 FileName:="Presentation Chart" 
 End If 
End With
```


## See also


#### Concepts


[Chart Object](chart-object-word.md)

