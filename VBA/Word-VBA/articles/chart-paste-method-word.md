---
title: Chart.Paste Method (Word)
keywords: vbawd10.chm79364167
f1_keywords:
- vbawd10.chm79364167
ms.prod: word
api_name:
- Word.Chart.Paste
ms.assetid: e159d28e-c2ff-9105-3b52-278fe55b078c
ms.date: 06/08/2017
---


# Chart.Paste Method (Word)

Pastes chart data from the Clipboard into the chart.


## Syntax

 _expression_ . **Paste**( **_Type_** )

 _expression_ A variable that represents a **[Chart](chart-object-word.md)** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Type_|Optional| **Variant**|Specifies the chart information to paste if a chart is on the Clipboard. Can be one of the following values: 

|-4104|Everything will be pasted. This is the default value.|
|-4122|Copied source formats are pasted.|
|-4123|Formulas are pasted.|
|

## Example

The following example activates the Microsoft Excel workbook associated with the first chart in the active document. If the Excel workbook has multiple windows, the example activates the first window. The example then copies the contents of cells B1 through B5 and pastes the cell contents into the chart.


```vb
With ActiveDocument.InlineShapes(1) 
 If .HasChart Then 
 .Chart.ChartData.Activate 
 .Chart.ChartData.Workbook. _ 
 Worksheets("Sheet1").Range("B1:B5").Copy 
 .Chart.Paste 
 End If 
End With 

```


## See also


#### Concepts


[Chart Object](chart-object-word.md)

