---
title: InvisibleApp.FormatResultEx Method (Visio)
keywords: vis_sdr.chm17551375
f1_keywords:
- vis_sdr.chm17551375
ms.prod: visio
api_name:
- Visio.InvisibleApp.FormatResultEx
ms.assetid: 8a6fe08f-06f2-f9d5-5699-aa40fec6bde8
ms.date: 06/08/2017
---


# InvisibleApp.FormatResultEx Method (Visio)

Formats a string or number into a string according to a format picture, using specified units for scaling and formatting. Optionally, for date or time strings, sets the language and calendar type of the string.


## Syntax

 _expression_ . **FormatResultEx**( **_StringOrNumber_** , **_UnitsIn_** , **_UnitsOut_** , **_Format_** , **_LangID_** , **_CalendarID_** , **_lpbstrRet_** )

 _expression_ A variable that represents an **InvisibleApp** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _StringOrNumber_|Required| **Variant**|String or number to be formatted; can be passed as a string, floating point number, or integer.|
| _UnitsIn_|Required| **Variant**|Measurement units to attribute to  _StringOrNumber_.|
| _UnitsOut_|Required| **Variant**|Measurement units to express the result in.|
| _Format_|Required| **String**|Picture of what the result string should look like.|
| _LangID_|Optional| **Long**|The language to use for the result string. |
| _CalendarID_|Optional| **Long**|he calendar to use for the result string. The default value is the Western calendar,  **visCalWestern** .|

### Return Value

String


## Remarks

If passed as a string,  _StringOrNumber_ might be the formula or prospective formula of a cell or the result or prospective result of a cell expressed as a string. The **FormatResultEx** method evaluates the string and formats the result. Because the string is being evaluated outside the context of being the formula of a particular cell, the **FormatResultEx** method returns an error if the string contains any cell references.

Possible values for  _StringOrNumber_ include:

1.7

3

"2.5"

"4.1 cm"

"12 ft - 17 in + (12 cm / SQRT(7))"

When  _UnitsIn_ is **visDate** , you can pass a numeric value to the DATETIME function in _StringOrNumber_. The integer portion of the value you pass should represent the number of days since December 30, 1899, and the decimal portion should represent the fraction of a day since midnight. For example, 38135.50 represents noon on May 28th, 2004.

The  _UnitsIn_ and _UnitsOut_ arguments can be strings such as "inches", "inch", "in.", or "i". Strings may be used for all supported Microsoft Visio units such as centimeters, meters, miles, and so on. You can also use any of the unit constants declared by the Visio type library in **[VisUnitCodes](visunitcodes-enumeration-visio.md)** . A list of valid units is also included in[About Units of Measure](http://msdn.microsoft.com/library/b6140312-b8e6-0cf2-9fe0-b14e800216bf%28Office.15%29.aspx).

If  _StringOrNumber_ is a string, _UnitsIn_ specifies how to interpret the evaluated result and is only used if the result is a scalar. For example, the expression " _4 * 5 cm_ " evaluates to 20 cm, which is not a scalar, so _UnitsIn_ is ignored. The expression " _4 * 5_ " evaluates to 20, which is a scalar and is interpreted using the specified _UnitsIn_.

The  _UnitsOut_ argument specifies the units in which the returned string should be expressed. If you want the results expressed in the same units as the evaluated expression, pass "NOCAST" or **visNoCast** .

 _Format_ is a string that specifies a template or picture of the string produced by the **FormatResultEx** method. For details, see the FORMAT function. A few of the possibilities are:

# : Output a single digit, but not if it is a leading or trailing 0.

0 : Output a single digit, even if it is a leading or trailing 0.

. : Decimal placeholder.

, : Thousands separator.

"text" or 'text' : Output enclosed text as is.

\c : Output the character c.

When  _UnitsIn_ is **visDate** , _Format_ should be one of the custom Microsoft Visio expanded-form date/time formats, which are of the form "{{ _date/time format picture_ }}". You can view these formats in the **Custom Format** box in the **Data Format** dialog box in Visio. (Select a shape, and then, on the **Insert** tab, click **Field**. In the  **Category** list, click **Date/Time**, and then click  **Data Format**.)

The  _LangID_ argument is optional. If you don't specify a value, Visio uses the current system language. If you pass a value, the _LangID_ argument should be one of the standard IDs used by Microsoft Windows to encode different language versions. For example, 1033 is the language ID for English (United States). To see a list of possible language IDs, search for "VERSIONINFO" in the Microsoft Platform SDK on MSDN.

The  _CalendarID_ argument should be one of the following values, which are declared in **VisCellVals** in the Visio type library. The default value is the Western calendar, **visCalWestern** .



|**Constant **|**Value **|**Description **|
|:-----|:-----|:-----|
| **visCalWestern**|0|Western|
| **visCalArabicHijri**|1|Arabic Hijiri|
| **visCalHebrewLunar**|2|Hebrew Lunar|
| **visCalChineseTaiwan**|3|Taiwan Calendar|
| **visCalJapaneseEmperor**|4|Japanese Emperor Reign|
| **visCalThaiBuddhism**|5|Thai Buddhist|
| **visCalKoreanDanki**|6|Korean Danki|
| **visCalSakaEra**|7|Saka Era|
| **visCalTranslitEnglish**|8|English transliterated |
| **visCalTranslitFrench**|9|French transliterated |

## Example

The following example shows how to use the  **FormatResultEx** property to format a date in Greek and display it as shape text.


```vb
Public Sub FormatResultEx_Example 
 
 Dim vsoShape As Visio.Shape 
 Dim strDate As String 
 
 Set vsoShape = ActivePage.DrawOval (3, 5, 5, 9) 
 strDate = Application.FormatResultEx (37663.50, visDate, "", "{{dd MMMM yyyy}}", 1032, 0) 
 
 vsoShape.Text = strDate 
 
End Sub
```


