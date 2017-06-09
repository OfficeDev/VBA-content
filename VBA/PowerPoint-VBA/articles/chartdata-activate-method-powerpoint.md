---
title: ChartData.Activate Method (PowerPoint)
keywords: vbapp10.chm689002
f1_keywords:
- vbapp10.chm689002
ms.prod: powerpoint
api_name:
- PowerPoint.ChartData.Activate
ms.assetid: 789651b8-334c-340a-e281-822f7129b76e
ms.date: 06/08/2017
---


# ChartData.Activate Method (PowerPoint)

Activates the first window of the workbook associated with the chart.


## Syntax

 _expression_. **Activate**

 _expression_ A variable that represents a **[ChartData](chartdata-object-powerpoint.md)** object.


## Remarks

If the chart is linked to a Microsoft Excel workbook, this method does not run any Auto_Activate or Auto_Deactivate macros that might be attached to the workbook (use the  **[RunAutoMacros](http://msdn.microsoft.com/library/85dfdadf-75e6-437d-fb7a-e17681a69b35%28Office.15%29.aspx)** method to run those macros).


 **Note**  You must call this method before referencing the  **[Workbook](chartdata-workbook-property-powerpoint.md)** property.


## Example




 **Note**  Although the following code applies to Microsoft Word, you can readily modify it to apply to PowerPoint.

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


[ChartData Object](chartdata-object-powerpoint.md)

