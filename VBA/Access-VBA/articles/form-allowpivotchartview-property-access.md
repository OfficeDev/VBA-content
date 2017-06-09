---
title: Form.AllowPivotChartView Property (Access)
keywords: vbaac10.chm13535
f1_keywords:
- vbaac10.chm13535
ms.prod: access
api_name:
- Access.Form.AllowPivotChartView
ms.assetid: 5585b530-d114-d07e-63cb-8d96dec458e8
ms.date: 06/08/2017
---


# Form.AllowPivotChartView Property (Access)

Returns or sets a  **Boolean** indicating whether the specified form may be viewed in PivotChart View. **True** if PivotChart View is allowed. Read/write.


## Syntax

 _expression_. **AllowPivotChartView**

 _expression_ A variable that represents a **Form** object.


## Remarks

Use the  **AllowDatasheetView**, **[AllowPivotChartView](form-allowpivotchartview-property-access.md)**, **AllowPivotChartView**, or **[AllowPivotTableView](form-allowpivottableview-property-access.md)** properties to control which views are allowed for a form.


## Example

The following example makes PivotChart View valid for the specified form and then opens the form in PivotChart View.


```vb
Forms(0).AllowPivotChartView = True 
DoCmd.OpenForm FormName:=Forms(0).Name, View:=acFormPivotChart 

```


## See also


#### Concepts


[Form Object](form-object-access.md)

