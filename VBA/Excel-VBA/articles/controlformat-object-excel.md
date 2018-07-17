---
title: ControlFormat Object (Excel)
keywords: vbaxl10.chm629072
f1_keywords:
- vbaxl10.chm629072
ms.prod: excel
api_name:
- Excel.ControlFormat
ms.assetid: fafc6e6b-641c-2179-0789-d86c2718b3c0
ms.date: 06/08/2017
---


# ControlFormat Object (Excel)

Contains Microsoft Excel control properties.


## Example

Use the  **[ControlFormat](shape-controlformat-property-excel.md)** property to return a **ControlFormat** object. The following example sets the fill range for a list box control on worksheet one.


 **Note**  If the shape isn't a control, the  **ControlFormat** property fails; and if the control isn't a list box, the **ListFillRange** property fails.


```vb
Worksheets(1).Shapes(1).ControlFormat.ListFillRange = "A1:A10"
```


## See also


#### Other resources


[Excel Object Model Reference](http://msdn.microsoft.com/library/11ea8598-8a20-92d5-f98b-0da04263bf2c%28Office.15%29.aspx)


