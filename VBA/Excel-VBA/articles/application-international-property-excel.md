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



| <strong>Index</strong>                   | <strong>Type</strong> | <strong>Meaning</strong>                                                          |
|:-----------------------------------------|:----------------------|:----------------------------------------------------------------------------------|
| <strong>xlLeftBrace</strong>             | String                | Character used instead of the left brace ({) in array literals.                   |
| <strong>xlLeftBracket</strong>           | String                | Character used instead of the left bracket ([) in R1C1-style relative references. |
| <strong>xlLowerCaseColumnLetter</strong> | String                | Lowercase column letter.                                                          |
| <strong>xlLowerCaseRowLetter</strong>    | String                | Lowercase row letter.                                                             |
| <strong>xlRightBrace</strong>            | String                | Character used instead of the right brace (}) in array literals.                  |
| <strong>xlRightBracket</strong>          | String                | Character used instead of the right bracket (]) in R1C1-style references.         |
| <strong>xlUpperCaseColumnLetter</strong> | String                | Uppercase column letter.                                                          |
| <strong>xlUpperCaseRowLetter</strong>    | String                | Uppercase row letter (for R1C1-style references).                                 |

 **Country/Region Settings**



| <strong>Index</strong>               | <strong>Type</strong> | <strong>Meaning</strong>                                     |
|:-------------------------------------|:----------------------|:-------------------------------------------------------------|
| <strong>xlCountryCode</strong>       | Long                  | Country/Region version of Microsoft Excel.                   |
| <strong>xlCountrySetting</strong>    | Long                  | Current country/region setting in the Windows Control Panel. |
| <strong>xlGeneralFormatName</strong> | String                | Name of the General number format.                           |

 **Currency**



| <strong>Index</strong>                   | <strong>Type</strong> | <strong>Meaning</strong>                                                                                                                                                                                                                                                                                                                                                                          |
|:-----------------------------------------|:----------------------|:--------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------|
| <strong>xlCurrencyBefore</strong>        | Boolean               | <strong>True</strong> if the currency symbol precedes the currency values; <strong>False</strong> if it follows them.                                                                                                                                                                                                                                                                             |
| <strong>xlCurrencyCode</strong>          | String                | Currency symbol.                                                                                                                                                                                                                                                                                                                                                                                  |
| <strong>xlCurrencyDigits</strong>        | Long                  | Number of decimal digits to be used in currency formats.                                                                                                                                                                                                                                                                                                                                          |
| <strong>xlCurrencyLeadingZeros</strong>  | Boolean               | <strong>True</strong> if leading zeros are displayed for zero currency values.                                                                                                                                                                                                                                                                                                                    |
| <strong>xlCurrencyMinusSign</strong>     | Boolean               | <strong>True</strong> if you?re using a minus sign for negative numbers; <strong>False</strong> if you?re using parentheses.                                                                                                                                                                                                                                                                      |
| <strong>xlCurrencyNegative</strong>      | Long                  | Currency format for negative currency values:0 = ( <em>symbol</em> x) or (x <em>symbol</em> )1 = - <em>symbol</em> x or -x <em>symbol_2 =  _symbol</em> -x or x- <em>symbol_3 =  _symbol</em> x- or x <em>symbol</em> -where  <em>symbol</em> is the currency symbol of the country or region. Note that the position of the currency symbol is determined by <strong>xlCurrencyBefore</strong> . |
| <strong>xlCurrencySpaceBefore</strong>   | Boolean               | <strong>True</strong> if a space is added before the currency symbol.                                                                                                                                                                                                                                                                                                                             |
| <strong>xlCurrencyTrailingZeros</strong> | Boolean               | <strong>True</strong> if trailing zeros are displayed for zero currency values.                                                                                                                                                                                                                                                                                                                   |
| <strong>xlNoncurrencyDigits</strong>     | Long                  | Number of decimal digits to be used in noncurrency formats.                                                                                                                                                                                                                                                                                                                                       |

 **Date and Time**



| <strong>Index</strong>              | <strong>Type</strong>    | <strong>Meaning</strong>                                                                                                                                    |
|:------------------------------------|:-------------------------|:------------------------------------------------------------------------------------------------------------------------------------------------------------|
| <strong>xl24HourClock</strong>      | <strong>Boolean</strong> | <strong>True</strong> if you?re using 24-hour time; <strong>False</strong> if you?re using 12-hour time.                                                    |
| <strong>xl4DigitYears</strong>      | <strong>Boolean</strong> | <strong>True</strong> if you?re using four-digit years; <strong>False</strong> if you?re using two-digit years.                                             |
| <strong>xlDateOrder</strong>        | <strong>Long</strong>    | Order of date elements:0 = month-day-year1 = day-month-year2 = year-month-day                                                                               |
| <strong>xlDateSeparator</strong>    | <strong>String</strong>  | Date separator ( <strong>/</strong> ).                                                                                                                      |
| <strong>xlDayCode</strong>          | <strong>String</strong>  | Day symbol (d).                                                                                                                                             |
| <strong>xlDayLeadingZero</strong>   | <strong>Boolean</strong> | <strong>True</strong> if a leading zero is displayed in days.                                                                                               |
| <strong>xlHourCode</strong>         | <strong>String</strong>  | Hour symbol (h).                                                                                                                                            |
| <strong>xlMDY</strong>              | <strong>Boolean</strong> | <strong>True</strong> if the date order is month-day-year for dates displayed in the long form; <strong>False</strong> if the date order is day-month-year. |
| <strong>xlMinuteCode</strong>       | <strong>String</strong>  | Minute symbol (m).                                                                                                                                          |
| <strong>xlMonthCode</strong>        | <strong>String</strong>  | Month symbol (m).                                                                                                                                           |
| <strong>xlMonthLeadingZero</strong> | <strong>Boolean</strong> | <strong>True</strong> if a leading zero is displayed in months (when months are displayed as numbers).                                                      |
| <strong>xlMonthNameChars</strong>   | <strong>Long</strong>    | Always returns three characters for backward compatibility. Abbreviated month names are read from Microsoft Windows and can be any length.                  |
| <strong>xlSecondCode</strong>       | <strong>String</strong>  | Second symbol (s).                                                                                                                                          |
| <strong>xlTimeSeparator</strong>    | <strong>String</strong>  | Time separator (:).                                                                                                                                         |
| <strong>xlTimeLeadingZero</strong>  | <strong>Boolean</strong> | <strong>True</strong> if a leading zero is displayed in times.                                                                                              |
| <strong>xlWeekdayNameChars</strong> | <strong>Long</strong>    | Always returns three characters for backward compatibility. Abbreviated weekday names are read from Microsoft Windows and can be any length.                |
| <strong>xlYearCode</strong>         | <strong>String</strong>  | Year symbol in number formats (y).                                                                                                                          |

 **Measurement Systems**



| <strong>Index</strong>                 | <strong>Type</strong>    | <strong>Meaning</strong>                                                                                                        |
|:---------------------------------------|:-------------------------|:--------------------------------------------------------------------------------------------------------------------------------|
| <strong>xlMetric</strong>              | <strong>Boolean</strong> | <strong>True</strong> if you?re using the metric system; <strong>False</strong> if you?re using the English measurement system. |
| <strong>xlNonEnglishFunctions</strong> | <strong>Boolean</strong> | <strong>True</strong> if you?re not displaying functions in English.                                                            |

 **Separators**



| <strong>Index</strong>                     | <strong>Type</strong>   | <strong>Meaning</strong>                                                                                       |
|:-------------------------------------------|:------------------------|:---------------------------------------------------------------------------------------------------------------|
| <strong>xlAlternateArraySeparator</strong> | <strong>String</strong> | Alternate array item separator to be used if the current array separator is the same as the decimal separator. |
| <strong>xlColumnSeparator</strong>         | <strong>String</strong> | Character used to separate columns in array literals.                                                          |
| <strong>xlDecimalSeparator</strong>        | <strong>String</strong> | Decimal separator.                                                                                             |
| <strong>xlListSeparator</strong>           | <strong>String</strong> | List separator.                                                                                                |
| <strong>xlRowSeparator</strong>            | <strong>String</strong> | Character used to separate rows in array literals.                                                             |
| <strong>xlThousandsSeparator</strong>      | <strong>String</strong> | Zero or thousands separator.                                                                                   |

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

