---
title: ChartData.Activate Method (Word)
keywords: vbawd10.chm190382081
f1_keywords:
- vbawd10.chm190382081
ms.prod: word
api_name:
- Word.ChartData.Activate
ms.assetid: 08f4a657-41c2-52ea-b31c-976549ace8c1
ms.date: 06/08/2017
---


# ChartData.Activate Method (Word)

Activates the first window of the workbook associated with the chart.


## Syntax

 _expression_ . **Activate**

 _expression_ A variable that represents a **[ChartData](chartdata-object-word.md)** object.


## Remarks

If the chart is linked to a Microsoft Excel workbook, this method does not run any Auto_Activate or Auto_Deactivate macros that might be attached to the workbook (use the  **[RunAutoMacros](http://msdn.microsoft.com/library/85dfdadf-75e6-437d-fb7a-e17681a69b35%28Office.15%29.aspx)** method to run those macros).


 **Note**  You must call this method before referencing the  **[Workbook](chartdata-workbook-property-word.md)** property.


## Example

The following example activates the Excel workbook associated with the first chart in the active document. If the Excel workbook has multiple windows, the example activates the first window. The example then copies the contents of cells B1 through B5 and pastes the cell contents into the chart.


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


[ChartData Object](chartdata-object-word.md)

