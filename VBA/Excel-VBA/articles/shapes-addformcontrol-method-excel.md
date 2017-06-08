---
title: Shapes.AddFormControl Method (Excel)
keywords: vbaxl10.chm638090
f1_keywords:
- vbaxl10.chm638090
ms.prod: excel
api_name:
- Excel.Shapes.AddFormControl
ms.assetid: c1654020-630c-b988-54f1-99a2f2a93e56
ms.date: 06/08/2017
---


# Shapes.AddFormControl Method (Excel)

Creates a Microsoft Excel control. Returns a  **[Shape](shape-object-excel.md)** object that represents the new control.


## Syntax

 _expression_ . **AddFormControl**( **_Type_** , **_Left_** , **_Top_** , **_Width_** , **_Height_** )

 _expression_ A variable that represents a **Shapes** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Type_|Required| **[XlFormControl](xlformcontrol-enumeration-excel.md)**|The Microsoft Excel control type. You cannot create an edit box on a worksheet.|
| _Left_|Required| **Long**|The initial coordinates of the new object (in points) relative to the upper-left corner of cell A1 on a worksheet or to the upper-left corner of a chart.|
| _Top_|Required| **Long**|The initial coordinates of the new object (in points) relative to the upper-left corner of cell A1 on a worksheet or to the upper-left corner of a chart.|
| _Width_|Required| **Long**|The initial size of the new object, in points.|
| _Height_|Required| **Long**|The initial size of the new object, in points.|

### Return Value

Shape


## Remarks

Use the  **[AddOLEObject](shapes-addoleobject-method-excel.md)** method or the **[Add](oleobjects-add-method-excel.md)** method of the **[OLEObjects](oleobjects-object-excel.md)** collection to create an ActiveX control.


## Example

This example adds a list box to worksheet one and sets the fill range for the list box.


```vb
With Worksheets(1) 
 Set lb = .Shapes.AddFormControl(xlListBox, 100, 10, 100, 100) 
 lb.ControlFormat.ListFillRange = "A1:A10" 
End With
```


## See also


#### Concepts


[Shapes Object](shapes-object-excel.md)

