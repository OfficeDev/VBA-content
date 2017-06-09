---
title: Cell.Result Property (Visio)
keywords: vis_sdr.chm10114195
f1_keywords:
- vis_sdr.chm10114195
ms.prod: visio
api_name:
- Visio.Cell.Result
ms.assetid: 5d97f8e7-0bb4-7334-8cf0-7fb3860fbc2b
ms.date: 06/08/2017
---


# Cell.Result Property (Visio)

Gets or sets a cell's value. Read/write.


## Syntax

 _expression_ . **Result**( **_UnitsNameOrCode_** )

 _expression_ A variable that represents a **Cell** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _UnitsNameOrCode_|Required| **Variant**|The units to use when retrieving or setting the cell's value.|

### Return Value

Double


## Remarks

Use the  **Result** property to set the value of an unguarded cell. If the cell's formula is protected with the GUARD function, the formula is not changed and an error is generated. If the cell contains only a text string, zero (0) is returned. If the string is invalid, an error is generated.

You can specify  _UnitsNameOrCode_ as an integer or a string value. For example, the following statements all set _UnitsNameOrCode_ to inches.

 _retVal_ = **Cell.Result** ( **visInches** )

 _retVal_ = **Cell.Result** (65)

 _retVal_ = **Cell.Result** ("in") where "in" can also be any of the alternate strings representing inches, such as "inch", "in.", or "intCounter".

For a complete list of valid unit strings along with corresponding Automation constants (integer values), see [About Units of Measure](http://msdn.microsoft.com/library/b6140312-b8e6-0cf2-9fe0-b14e800216bf%28Office.15%29.aspx).

Automation constants for representing units are declared by the Visio type library in member  **[VisUnitCodes ](visunitcodes-enumeration-visio.md)** .

To specify internal units, pass a zero-length string (""). Internal units are inches for distance and radians for angles. To specify implicit units, you must use the  **Formula** property.


## Example

This Microsoft Visual Basic for Applications (VBA) macro shows how to use the  **Result** property.


```vb
 
Public Sub Result_Example() 
 
 Dim vsoShape As Visio.Shape 
 Dim vsoCell As Visio.Cell 
 Dim intLocalCenterX As Double 
 
 'Draw a rectangle. 
 Set vsoShape = ActivePage.DrawRectangle(1, 5, 5, 1) 
 
 Set vsoCell = vsoShape.Cells("LocPinX") 
 intLocalCenterX = vsoCell.Result("cm") 
 Debug.Print intLocalCenterX 
 
 'You can also use the constants defined by the Visio type library. 
 intLocalCenterX = vsoCell.Result(visInches) 
 Debug.Print intLocalCenterX 
 
End Sub
```


