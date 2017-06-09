---
title: Font.Background Property (Excel)
keywords: vbaxl10.chm559073
f1_keywords:
- vbaxl10.chm559073
ms.prod: excel
api_name:
- Excel.Font.Background
ms.assetid: af7407c4-655a-5db7-abb2-6932675971d2
ms.date: 06/08/2017
---


# Font.Background Property (Excel)

Returns or sets the type of background for text used in charts. Read/write  **Variant** which is set to one of the constants of **[XlBackground](xlbackground-enumeration-excel.md)** .


## Syntax

 _expression_ . **Background**

 _expression_ A variable that represents a **Font** object.


## Remarks





|**XlBackground can be one of the following constants**|**Description of text background**|
|:-----|:-----|
| **xlBackgroundAutomatic**|Font background will automatically change the background area around the text to a color that best displays the chart text on the color applied to elements under the text|
| **xlBackgroundOpaque**|Font background will set the font bacground to black if the text color and fill color underneath the text are very close or the same color, such that the text would not appear|
| **xlBackgroundTransparent**|Font background is set to transparent so that text background does not change if the text color is close to the color underneath the text|

## Example

This example adds a chart title to embedded chart one on the first worksheet and then sets the font size and background type for the title. This example assumes a chart exists on the first worksheet.


```vb
Sub UseBackground() 
 
 With Worksheets(1).ChartObjects(1).Chart 
 .HasTitle = True 
 .ChartTitle.Text = "Rainfall Totals by Month" 
 With .ChartTitle.Font 
 .Size = 10 
 .Background = xlBackgroundTransparent 
 End With 
 End With 
 
End Sub
```


## See also


#### Concepts


[Font Object](font-object-excel.md)

