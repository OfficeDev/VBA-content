---
title: Application.International Property (Excel)
keywords: vbaxl10.chm133151
f1_keywords:
- vbaxl10.chm133151
ms.prod: excel
api_name:
- Excel.Application.International
ms.assetid: e3849e31-a808-256c-4a94-c75c9d674d66
ms.date: 06/08/2017
---


# Application.International Property (Excel)

Returns information about the current country/region and international settings. Read-only  **Variant** .


## Syntax

 _expression_ . **International**( **_Index_** )

 _expression_ A variable that represents an **Application** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Index_|Optional| **Variant**|The setting to be returned. Can be one of the  **XlApplicationInternational** constants listed in the following tables.|

## Remarks

 **Brackets and Braces**



|**Index**|**Type**|**Meaning**|
|:-----|:-----|:-----|
| **xlLeftBrace**|String|Character used instead of the left brace ({) in array literals.|
| **xlLeftBracket**|String|Character used instead of the left bracket ([) in R1C1-style relative references.|
| **xlLowerCaseColumnLetter**|String|Lowercase column letter.|
| **xlLowerCaseRowLetter**|String|Lowercase row letter.|
| **xlRightBrace**|String|Character used instead of the right brace (}) in array literals.|
| **xlRightBracket**|String|Character used instead of the right bracket (]) in R1C1-style references.|
| **xlUpperCaseColumnLetter**|String|Uppercase column letter.|
| **xlUpperCaseRowLetter**|String|Uppercase row letter (for R1C1-style references).|
 **Country/Region Settings**



|**Index**|**Type**|**Meaning**|
|:-----|:-----|:-----|
| **xlCountryCode**|Long|Country/Region version of Microsoft Excel.|
| **xlCountrySetting**|Long|Current country/region setting in the Windows Control Panel.|
| **xlGeneralFormatName**|String|Name of the General number format.|
 **Currency**



|**Index**|**Type**|**Meaning**|
|:-----|:-----|:-----|
| **xlCurrencyBefore**|Boolean| **True** if the currency symbol precedes the currency values; **False** if it follows them.|
| **xlCurrencyCode**|String|Currency symbol.|
| **xlCurrencyDigits**|Long|Number of decimal digits to be used in currency formats.|
| **xlCurrencyLeadingZeros**|Boolean| **True** if leading zeros are displayed for zero currency values.|
| **xlCurrencyMinusSign**|Boolean| **True** if you?re using a minus sign for negative numbers; **False** if you?re using parentheses.|
| **xlCurrencyNegative**|Long|Currency format for negative currency values:0 = ( _symbol_ x) or (x _symbol_ )1 = - _symbol_ x or -x _symbol_2 =  _symbol_ -x or x- _symbol_3 =  _symbol_ x- or x _symbol_ -where  _symbol_ is the currency symbol of the country or region. Note that the position of the currency symbol is determined by **xlCurrencyBefore** .|
| **xlCurrencySpaceBefore**|Boolean| **True** if a space is added before the currency symbol.|
| **xlCurrencyTrailingZeros**|Boolean| **True** if trailing zeros are displayed for zero currency values.|
| **xlNoncurrencyDigits**|Long|Number of decimal digits to be used in noncurrency formats.|
 **Date and Time**



|**Index**|**Type**|**Meaning**|
|:-----|:-----|:-----|
| **xl24HourClock**| **Boolean**| **True** if you?re using 24-hour time; **False** if you?re using 12-hour time.|
| **xl4DigitYears**| **Boolean**| **True** if you?re using four-digit years; **False** if you?re using two-digit years.|
| **xlDateOrder**| **Long**|Order of date elements:0 = month-day-year1 = day-month-year2 = year-month-day|
| **xlDateSeparator**| **String**|Date separator ( **/** ).|
| **xlDayCode**| **String**|Day symbol (d).|
| **xlDayLeadingZero**| **Boolean**| **True** if a leading zero is displayed in days.|
| **xlHourCode**| **String**|Hour symbol (h).|
| **xlMDY**| **Boolean**| **True** if the date order is month-day-year for dates displayed in the long form; **False** if the date order is day-month-year.|
| **xlMinuteCode**| **String**|Minute symbol (m).|
| **xlMonthCode**| **String**|Month symbol (m).|
| **xlMonthLeadingZero**| **Boolean**| **True** if a leading zero is displayed in months (when months are displayed as numbers).|
| **xlMonthNameChars**| **Long**|Always returns three characters for backward compatibility. Abbreviated month names are read from Microsoft Windows and can be any length.|
| **xlSecondCode**| **String**|Second symbol (s).|
| **xlTimeSeparator**| **String**|Time separator (:).|
| **xlTimeLeadingZero**| **Boolean**| **True** if a leading zero is displayed in times.|
| **xlWeekdayNameChars**| **Long**|Always returns three characters for backward compatibility. Abbreviated weekday names are read from Microsoft Windows and can be any length.|
| **xlYearCode**| **String**|Year symbol in number formats (y).|
 **Measurement Systems**



|**Index**|**Type**|**Meaning**|
|:-----|:-----|:-----|
| **xlMetric**| **Boolean**| **True** if you?re using the metric system; **False** if you?re using the English measurement system.|
| **xlNonEnglishFunctions**| **Boolean**| **True** if you?re not displaying functions in English.|
 **Separators**



|**Index**|**Type**|**Meaning**|
|:-----|:-----|:-----|
| **xlAlternateArraySeparator**| **String**|Alternate array item separator to be used if the current array separator is the same as the decimal separator.|
| **xlColumnSeparator**| **String**|Character used to separate columns in array literals.|
| **xlDecimalSeparator**| **String**|Decimal separator.|
| **xlListSeparator**| **String**|List separator.|
| **xlRowSeparator**| **String**|Character used to separate rows in array literals.|
| **xlThousandsSeparator**| **String**|Zero or thousands separator.|
Symbols, separators, and currency formats shown in the preceding table may differ from those used in your language or geographic location and may not be available to you, depending on the language support (U.S. English, for example) that you?ve selected or installed.


## Example

This example displays the international decimal separator.


```vb
MsgBox "The decimal separator is " &; _ 
 Application.International(xlDecimalSeparator)
```


## See also


#### Concepts


[Application Object](application-object-excel.md)

