---
title: Shape.FormControlType Property (Excel)
keywords: vbaxl10.chm636131
f1_keywords:
- vbaxl10.chm636131
ms.prod: excel
api_name:
- Excel.Shape.FormControlType
ms.assetid: a0f7d7e2-a5c0-fd71-bced-fe2866fc6d7f
ms.date: 06/08/2017
---


# Shape.FormControlType Property (Excel)

Returns the Microsoft Excel control type. Read-only  **[XlFormControl](xlformcontrol-enumeration-excel.md)** .


## Syntax

 _expression_ . **FormControlType**

 _expression_ A variable that represents a **Shape** object.


## Remarks

You cannot use this property with ActiveX controls (the  **[Type](shape-type-property-excel.md)** property for the **[Shape](shape-object-excel.md)** object must return **msoFormControl** ).


## Example

This example clears all the Microsoft Excel check boxes on worksheet one.


```vb
For Each s In Worksheets(1).Shapes 
 If s.Type = msoFormControl Then 
 If s.FormControlType = xlCheckBox Then _ 
 s.ControlFormat.Value = False 
 End If 
Next
```


## See also


#### Concepts


[Shape Object](shape-object-excel.md)

