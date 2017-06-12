---
title: Form.AllowFormView Property (Access)
keywords: vbaac10.chm13532
f1_keywords:
- vbaac10.chm13532
ms.prod: access
api_name:
- Access.Form.AllowFormView
ms.assetid: 15dc69fc-d4ba-c8e3-d047-71f96c32fe02
ms.date: 06/08/2017
---


# Form.AllowFormView Property (Access)

Returns or sets a  **Boolean** indicating whether the specified form may be viewed in Form View. **True** if Form View is allowed. Read/write.


## Syntax

 _expression_. **AllowFormView**

 _expression_ A variable that represents a **Form** object.


## Remarks

Use the  **AllowDatasheetView**, **AllowFormView**, **[AllowPivotChartView](form-allowpivotchartview-property-access.md)**, or **[AllowPivotTableView](form-allowpivottableview-property-access.md)** properties to control which views are allowed for a form.


## Example

The following example makes Form View valid for the specified form and then opens the form in Form View.


```vb
Forms(0).AllowFormView = True 
DoCmd.OpenForm FormName:=Forms(0).Name, View:=acNormal
```


## See also


#### Concepts


[Form Object](form-object-access.md)

