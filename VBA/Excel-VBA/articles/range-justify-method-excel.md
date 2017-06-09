---
title: Range.Justify Method (Excel)
keywords: vbaxl10.chm144152
f1_keywords:
- vbaxl10.chm144152
ms.prod: excel
api_name:
- Excel.Range.Justify
ms.assetid: f8b4d48b-8cbb-977a-fd44-d354661182d2
ms.date: 06/08/2017
---


# Range.Justify Method (Excel)

Rearranges the text in a range so that it fills the range evenly.


## Syntax

 _expression_ . **Justify**

 _expression_ A variable that represents a **Range** object.


### Return Value

Variant


## Remarks

If the range isn't large enough, Microsoft Excel displays a message telling you that text will extend below the range. If you click the  **OK** button, justified text will replace the contents in cells that extend beyond the selected range. To prevent this message from appearing, set the **[DisplayAlerts](application-displayalerts-property-excel.md)** property to **False** . After you set this property, text will always replace the contents in cells below the range.


## Example

This example justifies the text in cell A1 on Sheet1.


```vb
Worksheets("Sheet1").Range("A1").Justify
```


## See also


#### Concepts


[Range Object](range-object-excel.md)

