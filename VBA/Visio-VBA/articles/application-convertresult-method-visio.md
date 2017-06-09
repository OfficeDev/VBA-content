---
title: Application.ConvertResult Method (Visio)
keywords: vis_sdr.chm10016135
f1_keywords:
- vis_sdr.chm10016135
ms.prod: visio
api_name:
- Visio.Application.ConvertResult
ms.assetid: b326c9cf-a7f3-33d7-1b29-8d1360301a9d
ms.date: 06/08/2017
---


# Application.ConvertResult Method (Visio)

Converts a string or number into an equivalent number in different measurement units.


## Syntax

 _expression_ . **ConvertResult**( **_StringOrNumber_** , **_UnitsIn_** , **_UnitsOut_** )

 _expression_ A variable that represents an **Application** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _StringOrNumber_|Required| **Variant**|String or number to be converted; can be a string, floating point number, or integer.|
| _UnitsIn_|Required| **Variant**|Measurement units to attribute to  _StringOrNumber_.|
| _UnitsOut_|Required| **Variant**|Measurement units to express the result in.|

### Return Value

Double


## Remarks

If passed as a string,  _StringOrNumber_ might be the formula or prospective formula of a cell or the result or prospective result of a cell expressed as a string. The **ConvertResult** method evaluates the string and converts the result into the units designated by _UnitsOut_. The  **ConvertResult** method returns an error if the string contains any cell references.

Possible values for  _StringOrNumber_ include:

1.7

3

"2.5"

"4.1 cm"

"12 ft - 17 in + (12 cm / SQRT(7))"

The  _UnitsIn_ and _UnitsOut_ arguments can be strings such as "inches", "inch", "in.", or "i". Strings may be used for all supported Microsoft Office Visio units such as centimeters, meters, miles, and so on. You can also use any of the units constants declared by the Visio type library in **VisUnitCodes** . A list of valid units is also listed in[About Units of Measure](http://msdn.microsoft.com/library/b6140312-b8e6-0cf2-9fe0-b14e800216bf%28Office.15%29.aspx).

If  _StringOrNumber_ is a floating point number or integer, _UnitsIn_ declares what unit of measure the **ConvertResult** method should construe the number to be. Pass "" to indicate internal Visio units.

If  _StringOrNumber_ is a string, _UnitsIn_ specifies how to interpret the evaluated result and is only used if the result is a scalar. For example, the expression "4 * 5 cm" evaluates to 20 cm, which is not a scalar, so _UnitsIn_ is ignored. The expression "4 * 5" evaluates to 20 which is a scalar and is interpreted using the specified _UnitsIn_.

The  _UnitsOut_ argument specifies in what units the returned number should be expressed. If you want the results expressed in the same units as the evaluated expression, pass "NOCAST" or **visNoCast** .

Examples where string is specified:




```
 
Debug.Print vsoApplication.ConvertResult("0.5 * 2", "ft", "ft") >>> 1.0 
Debug.Print vsoApplication.ConvertResult("0.5 * 2", "ft", "in") >>> 12.0 
Debug.Print vsoApplication.ConvertResult("1 cm", "ft", "in") >>> 0.39 
Debug.Print vsoApplication.ConvertResult("1 cm", "ft", "NOCAST") >>> 1.0 
Debug.Print vsoApplication.ConvertResult("1 cm", "ft", "") >>> 0.39 
Debug.Print vsoApplication.ConvertResult("1 cm", "ft", "bz") >>> exception: Bad measurement unit. 

```

Examples where number is specified:




```
 
Debug.Print vsoApplication.ConvertResult(1, "ft", "ft") >>> 1 
Debug.Print vsoApplication.ConvertResult(1, "ft", "in") >>> 12 
Debug.Print vsoApplication.ConvertResult(1.0, "in", "ft") >>> 8.33333333333333E-02 
Debug.Print vsoApplication.ConvertResult(1.0, visFeet, "") >>> 12 
Debug.Print vsoApplication.ConvertResult(1, "bz", "in") >>> exception: Bad measurement unit. 

```


## Example

The following macro shows how to use the  **ConvertResult** method to report the distance between two shapes in centimeters, feet, yards, and miles. To run this macro, you must have two shapes selected on your page.


```vb
Sub ConvertResult_Example() 
 
 Dim vsoApplication As Visio.Application 
 Dim vsoWindow As Visio.Window 
 Dim vsoSelection As Visio.Selection 
 Dim vsoShape1 As Visio.Shape 
 Dim vsoShape2 As Visio.Shape 
 
 Dim dblPinX1 As Double 
 Dim dblPinY1 As Double 
 Dim dblPinX2 As Double 
 Dim dblPinY2 As Double 
 Dim dblPinX1in As Double 
 Dim dblPinY1in As Double 
 Dim dblPinX2in As Double 
 Dim dblPinY2in As Double 
 Dim lngCount As Long 
 Dim dblDistance As Double 
 Dim dblDistanceX As Double 
 Dim dblDistanceY As Double 
 Dim dblResult(4) As Double 
 Dim strUnit As String 
 Set vsoApplication = Visio.Application 
 Set vsoWindow = vsoApplication.ActiveWindow 
 
 'Drawing page must be active window 
 If vsoWindow.Type = 1 Then 
 Set vsoSelection = vsoWindow.Selection 
 lngCount = vsoSelection.Count 
 
 'Exactly two shapes should be selected 
 If lngCount <> 2 Then 
 MsgBox "A total of " &; lngCount &; " shapes are " _ 
 &; "selected. Please select two shapes and try " _ 
 &; "again", 0 
 Else 
 Set vsoShape1 = vsoSelection.Item(1) 
 Set vsoShape2 = vsoSelection.Item(2) 
 
 'Pass the Visio Automation constant for inches (visInches, which is defined as 65) to the Result method to force units to inches 
 dblPinX1in = vsoShape1.Cells("PinX").Result(65) 
 dblPinY1in = vsoShape1.Cells("PinY").Result(65) 
 dblPinX2in = vsoShape2.Cells("PinX").Result(65) 
 dblPinY2in = vsoShape2.Cells("PinY").Result(65) 
 dblDistance = Sqr((dblPinX2in - dblPinX1in) ^ 2 + _ 
 (dblPinY2in - dblPinY1in) ^ 2) 
 
 'Convert distances from inches to centimeters, feet, yards, and miles 
 dblResult(1) = vsoApplication.ConvertResult(dblDistance, "in", "cm") 
 dblResult(2) = vsoApplication.ConvertResult(dblDistance, "in", "ft") 
 dblResult(3) = vsoApplication.ConvertResult(dblDistance, "in", "yd") 
 dblResult(4) = vsoApplication.ConvertResult(dblDistance, "in", "mi") 
 
 'Display results 
 MsgBox dblResult(1) &; " centimeters; " &; dblResult(2) &; " feet; " &; _ 
 dblResult(3) &; " ;yards; " &; dblResult(4) &; " miles ", 0 
 
 End If 
 
 Else 
 MsgBox "The drawing page must be active.", 0 
 
 End If 
 
End Sub
```


