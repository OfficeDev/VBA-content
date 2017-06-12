---
title: Form.AllowPivotTableView Property (Access)
keywords: vbaac10.chm13534,vbaac10.chm5540
f1_keywords:
- vbaac10.chm13534,vbaac10.chm5540
ms.prod: access
api_name:
- Access.Form.AllowPivotTableView
ms.assetid: 42bad4b4-7de1-f144-9482-2e114fc5cc4b
ms.date: 06/08/2017
---


# Form.AllowPivotTableView Property (Access)

Returns or sets a  **Boolean** indicating whether the specified form may be viewed in PivotTable View. **True** if PivotTable View is allowed. Read/write.


## Syntax

 _expression_. **AllowPivotTableView**

 _expression_ A variable that represents a **Form** object.


## Remarks

Use the  **[AllowDatasheetView](form-allowdatasheetview-property-access.md)**, **[AllowFormView](form-allowformview-property-access.md)**, **AllowPivotChartView**, or **AllowPivotTableView** properties to control which views are allowed for a form.


## Example

The following example makes PivotTable View valid for the specified form and then opens the form in PivotTable View.


```vb
Forms(0).AllowPivotTableView = True 
DoCmd.OpenForm FormName:=Forms(0).Name, View:=acFormPivotTable 

```


## See also


#### Concepts


[Form Object](form-object-access.md)

