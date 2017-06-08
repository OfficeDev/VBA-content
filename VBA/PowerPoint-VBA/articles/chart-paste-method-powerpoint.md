---
title: Chart.Paste Method (PowerPoint)
keywords: vbapp10.chm684004
f1_keywords:
- vbapp10.chm684004
ms.prod: powerpoint
api_name:
- PowerPoint.Chart.Paste
ms.assetid: b94ee91d-5b7c-c0b3-340d-b8eba4f3710f
ms.date: 06/08/2017
---


# Chart.Paste Method (PowerPoint)

Pastes chart data from the Clipboard into the chart.


## Syntax

 _expression_. **Paste**( **_Type_** )

 _expression_ A variable that represents a **[Chart](chart-object-powerpoint.md)** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Type_|Optional|**Variant**|Specifies the chart information to paste if a chart is on the Clipboard. Can be one of the following values: 
|||
|:-----|:-----|
|-4104|Everything will be pasted. This is the default value.|
|-4122|Copied source formats are pasted.|
|-4123|Formulas are pasted.|
|

## Example




 **Note**  Although the following code applies to Microsoft Word, you can readily modify it to apply to PowerPoint.

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


[Chart Object](chart-object-powerpoint.md)

