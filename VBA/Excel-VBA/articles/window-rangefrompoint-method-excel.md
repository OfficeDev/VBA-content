---
title: Window.RangeFromPoint Method (Excel)
keywords: vbaxl10.chm356131
f1_keywords:
- vbaxl10.chm356131
ms.prod: excel
api_name:
- Excel.Window.RangeFromPoint
ms.assetid: ece6172d-013d-5175-55e3-4968947d9e4e
ms.date: 06/08/2017
---


# Window.RangeFromPoint Method (Excel)

Returns the  **[Shape](shape-object-excel.md)** or **[Range](range-object-excel.md)** object that is positioned at the specified pair of screen coordinates. If there isn?t a shape located at the specified coordinates, this method returns **Nothing** .


## Syntax

 _expression_ . **RangeFromPoint**( **_x_** , **_y_** )

 _expression_ A variable that represents a **Window** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _x_|Required| **Long**|The value (in pixels) that represents the horizontal distance from the left edge of the screen, starting at the top.|
| _y_|Required| **Long**|The value (in pixels) that represents the vertical distance from the top of the screen, starting on the left.|

### Return Value

Object


## Example

This example returns the alternative text for the shape immediately below the mouse pointer if the shape is a chart, line, or picture.


```vb
Private Function AltText(ByVal intMouseX As Integer, _ 
 ByVal intMouseY as Integer) As String 
 Set objShape = ActiveWindow.RangeFromPoint _ 
 (x:=intMouseX, y:=intMouseY) 
 If Not objShape Is Nothing Then 
 With objShape 
 Select Case .Type 
 Case msoChart, msoLine, msoPicture: 
 AltText = .AlternativeText 
 Case Else: 
 AltText = "" 
 End Select 
 End With 
 Else 
 AltText = "" 
 End If 
End Function
```


## See also


#### Concepts


[Window Object](window-object-excel.md)

