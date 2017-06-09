---
title: Application.FormatResult Method (Visio)
keywords: vis_sdr.chm10016300
f1_keywords:
- vis_sdr.chm10016300
ms.prod: visio
api_name:
- Visio.Application.FormatResult
ms.assetid: 1b2178ab-e2ed-b618-ad2a-d18196f50be2
ms.date: 06/08/2017
---


# Application.FormatResult Method (Visio)

Formats a string or number into a string according to a format picture. Uses specified units for scaling and formatting.


## Syntax

 _expression_ . **FormatResult**( **_StringOrNumber_** , **_UnitsIn_** , **_UnitsOut_** , **_Format_** )

 _expression_ A variable that represents an **Application** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _StringOrNumber_|Required| **Variant**|String or number to be formatted; can be passed as a string, floating point number, or integer.|
| _UnitsIn_|Required| **Variant**|Measurement units to attribute to  _StringOrNumber_.|
| _UnitsOut_|Required| **Variant**|Measurement units to express the result in.|
| _Format_|Required| **String**|Picture of what the result string should look like.|

### Return Value

String


## Remarks

If passed as a string,  _StringOrNumber_ might be the formula or prospective formula of a cell or the result or prospective result of a cell expressed as a string. The **FormatResult** method evaluates the string and formats the result. Because the string is being evaluated outside the context of being the formula of a particular cell, the **FormatResult** method returns an error if the string contains any cell references.

Possible values for  _StringOrNumber_ include:

1.7

3

"2.5"

"4.1 cm"

"12 ft - 17 in. + (12 cm / SQRT(7))"

The  _UnitsIn_ and _UnitsOut_ arguments can be strings such as "inches", "inch", "in.", or "i". Strings may be used for all supported Microsoft Office Visio units such as centimeters, meters, miles, and so on. You can also use any of the unit constants declared by the Visio type library in **[VisUnitCodes](visunitcodes-enumeration-visio.md)** . A list of valid units is also included in[About Units of Measure](http://msdn.microsoft.com/library/b6140312-b8e6-0cf2-9fe0-b14e800216bf%28Office.15%29.aspx).

If  _StringOrNumber_ is a string, _UnitsIn_ specifies how to interpret the evaluated result and is only used if the result is a scalar. For example, the expression " _4 * 5 cm_ " evaluates to 20 cm, which is not a scalar, so _UnitsIn_ is ignored. The expression " _4 * 5_ " evaluates to 20, which is a scalar and is interpreted using the specified _UnitsIn_ .

The  _UnitsOut_ argument specifies the units in which the returned string should be expressed. If you want the results expressed in the same units as the evaluated expression, pass "NOCAST" or **visNoCast** .

 _Format_ is a string that specifies a template or picture of the string produced by the **FormatResult** method. For details, see the FORMAT function. A few of the possibilities are:

# : Output a single digit, but not if it's a leading or trailing 0.

0 : Output a single digit, even if it is a leading or trailing 0.

. : Decimal placeholder.

, : Thousands separator.

"text" or 'text' : Output enclosed text as is.

\c : Output the character c.


## Example

Where a string is specified


```vb
' Prints 1.00 
Debug.Print Application.FormatResult("0.5 * 2", "ft", "ft", "#.00 u") 
 
' Prints 12.00 in. 
Debug.Print Application.FormatResult("0.5 * 2", "ft", "in", "#.00 u") 
 
' Prints .39 in. 
Debug.Print Application.FormatResult("1 cm", "ft", "in", "#.00 u") 
 
' Prints 1.00 cm. 
Debug.Print Application.FormatResult("1 cm", "ft", "NOCAST", "#.00 u") 
 
' Prints 0.39 
Debug.Print Application.FormatResult("1 cm", "ft", "", "0.00 u") 
 
' Prints 1858.06 sq. cm. 
Debug.Print Application.FormatResult("1 sq. ft. * 2", "in^2", "cm^2", "0.00 u") 
 
' Throws an exception because of bad measurement unit ("bz") 
Debug.Print Application.FormatResult("1 cm", "ft", "bz", "#.00 u") 
```

Where a number is specified




```vb
' Prints 1.00 
Debug.Print Application.FormatResult(1, "ft", "ft", "#.00 u") 
 
' Prints 12.00 in. 
Debug.Print Application.FormatResult(1, "ft", "in", "#.00 u") 
 
' Prints .08 ft. 
Debug.Print Application.FormatResult(1.0, "in", "ft", "#.00 u") 
 
' Prints 12.00 
Debug.Print Application.FormatResult(1.0, visFeet, "", "#.00 u") 
 
' Throws an exception because of bad measurement unit ("bz") 
Debug.Print Application.FormatResult(1, "bz", "in", "#.00 u") 

```

The following macro shows how to use the  **FormatResult** method to convert a value from centimeters to inches and display the result in a message box.




```vb
 
Public Sub FormatResult_Example() 
 
 Dim strOldValue As String 
 Dim strNewValue As String 
 
 'Set old value. 
 strOldValue = "1 cm" 
 
 'Format value. 
 strNewValue = Application.FormatResult _ 
 (strOldValue, "ft", "in", "#.00 u") 
 
 'Display new value. 
 MsgBox (strNewValue) 
 
End Sub
```


