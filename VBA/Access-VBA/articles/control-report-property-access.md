---
title: Control.Report Property (Access)
keywords: vbaac10.chm10140
f1_keywords:
- vbaac10.chm10140
ms.prod: access
api_name:
- Access.Control.Report
ms.assetid: 1c1f4703-bda7-de97-eb13-830238a5170a
ms.date: 06/08/2017
---


# Control.Report Property (Access)

You can use the  **Report** property to refer to a report or to refer to the report associated with a subreport control. Read-only **Report**.


## Syntax

 _expression_. **Report**

 _expression_ A variable that represents a **Control** object.


## Remarks

This property is typically used to refer to the report contained in a subreport control.


 **Note**  When you use the  **[Reports](reports-object-access.md)** collection, you must specify the name of the report.


## Example

The following example uses the  **Report** property to refer to a control on a subreport.


```vb
Dim curTotalSales As Currency 
 
curTotalSales = Reports!Sales!Employees.Report!TotalSales
```


## See also


#### Concepts


[Control Object](control-object-access.md)

