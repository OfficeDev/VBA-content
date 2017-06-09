---
title: Selection.Move Method (Visio)
keywords: vis_sdr.chm11151355
f1_keywords:
- vis_sdr.chm11151355
ms.prod: visio
api_name:
- Visio.Selection.Move
ms.assetid: 12e60f50-f06d-45bb-b79d-db2e0d767461
ms.date: 06/08/2017
---


# Selection.Move Method (Visio)

Moves a selection a specified distance.


## Syntax

 _expression_ . **Move**( **_dx_** , **_dy_** , **_UnitsNameOrCode_** )

 _expression_ A variable that represents a **Selection** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _dx_|Required| **Double**|Specifies the amount to move in the x-direction.|
| _dy_|Required| **Double**|Specifies the amount to move in the y-direction.|
| _UnitsNameOrCode_|Optional| **Variant**|Specifies the units to use for  _dx_ and _dy_. See Remarks for possible values. The default is inches.|

### Return Value

Nothing


## Remarks

You can specify  _UnitsNameOrCode_ as an integer (a member of **[VisUnitCodes](visunitcodes-enumeration-visio.md)** ) or a string value such as "inches". If the string is invalid or the unit code is inappropriate (nontextual), an error is generated.

For a complete list of valid unit strings along with corresponding Automation constants (integer values), see [About units of measure](http://msdn.microsoft.com/library/b6140312-b8e6-0cf2-9fe0-b14e800216bf%28Office.15%29.aspx).


## Example

This Microsoft Visual Basic for Applications (VBA) macro shows how to use the  **Move** method to move a selection by a specified amount.


```vb
Public Sub Move_Example() 
 
 Dim vsoShape1 As Visio.Shape 
 Dim vsoShape2 As Visio.Shape 
 
 Set vsoShape1 = Application.ActiveWindow.Page.DrawRectangle(1, 9, 3, 7) 
 Set vsoShape2 = Application.ActiveWindow.Page.DrawRectangle(3, 6, 5, 5) 
 
 ActiveWindow.DeselectAll 
 
 ActiveWindow.Select vsoShape1, visSelect 
 ActiveWindow.Select vsoShape2, visSelect 
 Application.ActiveWindow.Selection.Move 2, 2 
 
End Sub
```


