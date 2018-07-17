---
title: Selection.Rotate Method (Visio)
keywords: vis_sdr.chm11151330
f1_keywords:
- vis_sdr.chm11151330
ms.prod: visio
api_name:
- Visio.Selection.Rotate
ms.assetid: 3c0a1a4d-a172-131a-9fb4-d215a5b9b2af
ms.date: 06/08/2017
---


# Selection.Rotate Method (Visio)

Rotates selected shapes either as a group or individually about their pins.


## Syntax

 _expression_ . **Rotate**( **_Angle_** , **_AngleUnitsNameOrCode_** , **_BlastGuards_** , **_RotationType_** , **_PinX_** , **_PinY_** , **_PinUnitsNameOrCode_** )

 _expression_ A variable that represents a **Selection** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Angle_|Required| **Double**|Specifies the angle to rotate the selection. See Remarks for possible values.|
| _AngleUnitsNameOrCode_|Optional| **Variant**|Specifies the units to use for  _Angle_. See Remarks for possible values. The default is degrees.|
| _BlastGuards_|Optional| **Boolean**| **True** to override formulas in the ShapeSheet of any of the selected shapes to which the GUARD function has been applied; **False** to leave guarded formulas unchanged. The default is **False** .|
| _RotationType_|Optional| **VisRotationTypes**|Specifes how the selection is to be rotated. See Remarks for possible values.|
| _PinX_|Optional| **Double**|When  _RotationType_ is **visRotateSelectionWithPin** , specifies the X-position of the pin about which the selection is to be rotated.|
| _PinY_|Optional| **Double**| When _RotationType_ is **visRotateSelectionWithPin** , specifies the Y-position of the pin about which the selection is to be rotated.|
| _PinUnitsNameOrCode_|Optional| **Variant**|Specifies the units to use for  _PinX_ and _PinY_. See Remarks for possible values. The default is inches.|

### Return Value

Nothing


## Remarks

The following possible values for  _RotationType_ are declared in **VisRotationTypes** in the Visio type library.



|**Constant**|**Value**|**Description**|
|:-----|:-----|:-----|
| **visRotateSelectionWithPin**|1|Rotates the selection around a pin.|
| **visRotateSelection**|0|Rotates the selection relative to the center of the selection.|
| **visRotateShapes**|2|Rotates the selected shapes around their pins relative to their current angle.|
Passing  **True** for the optional _BlastGuards_ argument overrides formulas in the ShapeSheet of any of the selected shapes to which the GUARD function has been applied.

The default value for  _RotationType_ is **visRotateSelection** .

You can specify  _AngleUnitsNameOrCode_ or _PinUnitsNameOrCode_ as an integer (a member of **[VisUnitCodes](visunitcodes-enumeration-visio.md)** ) or a string value such as "radians" or "inches". If the string is invalid or the unit code is inappropriate (nontextual), an error is generated.

For a complete list of valid unit strings along with corresponding Automation constants (integer values), see [About units of measure](http://msdn.microsoft.com/library/b6140312-b8e6-0cf2-9fe0-b14e800216bf%28Office.15%29.aspx).


## Example

This Microsoft Visual Basic for Applications (VBA) macro shows how to use the  **Rotate** method to rotate a selection 45 degrees relative to the center of the selection.


```vb
Public Sub Rotate_Example() 
 
 Dim vsoShape1 As Visio.Shape 
 Dim vsoShape2 As Visio.Shape 
 
 Set vsoShape1 = Application.ActiveWindow.Page.DrawRectangle(1, 9, 3, 7) 
 Set vsoShape2 = Application.ActiveWindow.Page.DrawRectangle(3, 6, 5, 5) 
 
 ActiveWindow.DeselectAll 
 
 ActiveWindow.Select vsoShape1, visSelect 
 ActiveWindow.Select vsoShape2, visSelect 
 
 Application.ActiveWindow.Selection.Rotate 45, visDegrees 
 
End Sub
```


