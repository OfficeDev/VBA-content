---
title: Form.AllowDatasheetView Property (Access)
keywords: vbaac10.chm13533
f1_keywords:
- vbaac10.chm13533
ms.prod: access
api_name:
- Access.Form.AllowDatasheetView
ms.assetid: 81796b90-94dd-cd27-3613-a2050e2bce21
ms.date: 06/08/2017
---


# Form.AllowDatasheetView Property (Access)

Returns or sets a  **Boolean** indicating whether the specified form may be viewed in Datasheet View. **True** if Datasheet View is allowed. Read/write.


## Syntax

 _expression_. **AllowDatasheetView**

 _expression_ A variable that represents a **Form** object.


## Remarks

Use the  **AllowDatasheetView**, **AllowFormView**, **[AllowPivotChartView](form-allowpivotchartview-property-access.md)**, or **[AllowPivotTableView](form-allowpivottableview-property-access.md)** properties to control which views are allowed for a form.


## Example

The following example makes Datasheet View valid for the specified form and then opens the form in Datasheet View.


```vb
Forms(0).AllowDatasheetView = True 
DoCmd.OpenForm FormName:=Forms(0).Name, View:=acFormDS 

```


## See also


#### Concepts


[Form Object](form-object-access.md)

