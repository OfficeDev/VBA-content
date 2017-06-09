---
title: Format Function (Visual Basic for Applications)
keywords: vblr6.chm1008925
f1_keywords:
- vblr6.chm1008925
ms.prod: office
ms.assetid: 67f60abf-0c77-49ec-924f-74ae6eb96ea8
ms.date: 06/08/2017
---


# Format Function (Visual Basic for Applications)



 **Description**
Returns a  **Variant (String)** containing an [expression](vbe-glossary.md) formatted according to instructions contained in a format expression.
 **Syntax**
 **Format(**_expression_ [ **,**_format_ [ **,**_firstdayofweek_ [ **,**_firstweekofyear_ ]]] **)**
The  **Format** function syntax has these parts:


|**Part**|**Description**|
|:-----|:-----|
| _expression_|Required. Any valid expression.|
| _format_|Optional. A valid named or user-defined format expression.|
| _firstdayofweek_|Optional. A [constant](vbe-glossary.md) that specifies the first day of the week.|
| _firstweekofyear_|Optional. A constant that specifies the first week of the year.|
 **Settings**
The  _firstdayofweek_[argument](vbe-glossary.md) has these settings:


|**Constant**|**Value**|**Description**|
|:-----|:-----|:-----|
|**vbUseSystem**|0|Use NLS API setting.|
|**vbSunday**|1|Sunday (default)|
|**vbMonday**|2|Monday|
|**vbTuesday**|3|Tuesday|
|**vbWednesday**|4|Wednesday|
|**vbThursday**|5|Thursday|
|**vbFriday**|6|Friday|
|**vbSaturday**|7|Saturday|
The  _firstweekofyear_ argument has these settings:


|**Constant**|**Value**|**Description**|
|:-----|:-----|:-----|
|**vbUseSystem**|0|Use NLS API setting.|
|**vbFirstJan1**|1|Start with week in which January 1 occurs (default).|
|**vbFirstFourDays**|2|Start with the first week that has at least four days in the year.|
|**vbFirstFullWeek**|3|Start with the first full week of the year.|
 **Remarks**


|**To Format**|**Do This**|
|:-----|:-----|
|Numbers|Use predefined named numeric formats or create user-defined numeric formats.|
|Dates and times|Use predefined named date/time formats or create user-defined date/time formats.|
|Date and time serial numbers|Use date and time formats or numeric formats.|
|Strings|Create your own user-defined string formats.|
If you try to format a number without specifying  _format_, **Format** provides functionality similar to the **Str** function, although it is internationally aware. However, positive numbers formatted as strings using **Format** don't include a leading space reserved for the sign of the value; those converted using **Str** retain the leading space.
If you are formatting a non-localized numeric string, you should use a user-defined numeric format to ensure that you get the look you want.

 **Note**  If the  **Calendar** property setting is `Gregorian` and _format_ specifies date formatting, the supplied _expression_ must be `Gregorian`. If the Visual Basic  **Calendar** property setting is `Hijri`, the supplied  _expression_ must be `Hijri`.

If the calendar is Gregorian, the meaning of  _format_ expression symbols is unchanged. If the calendar is Hijri, all date format symbols (for example, _dddd_, _mmmm_, _yyyy_ ) have the same meaning but apply to the Hijri calendar. Format symbols remain in English; symbols that result in text display (for example, AM and PM) display the string (English or Arabic) associated with that symbol. The range of certain symbols changes when the calendar is Hijri.

**Date Symbols**

|**Symbol**|**Range**|
|:-----|:-----|
| _d_|1-31 (Day of month, with no leading zero).|
| _dd_|01-31 (Day of month, with a leading zero).|
| _w_ |1-7 (Day of week, starting with Sunday = 1)
| _ww_|1-53 (Week of year, with no leading zero; Week 1 starts on Jan 1).|
| _mmm_|Displays abbreviated month names (Hijri month names have no abbreviations).|
| _mmmm_|Displays full month names.|
| _y_|1-366 (Day of year)|
| _yy_ |00-99 (Last two digits of year)|
| _yyyy_ |100-9666 (Three- or Four-digit year)|

**Time Symbols**

|**Symbol**|**Range**
|:-----|:-----|
| _h_ |0-23 (1-12 with "Am" or "Pm" appended) (Hour of day, with no leading zero)|
| _hh_ |0-23 (01-12 with "Am" or "Pm" appended) (Hour of day, with no leading zero)|
| _n_ |0-59 (Minute of hour, with no leading zero)|
| _nn_ |0-59 (Minute of hour, with a leading zero)|
| _s_ |0-59 (Second of minute, with no leading zero)|
| _ss_ |0-59 (Second of minute, with a leading zero)|


 **Example**
 
This example shows various uses of the  **Format** function to format values using both named formats and user-defined formats. For the date separator ( **/** ), time separator ( **:** ), and AM/ PM literal, the actual formatted output displayed by your system depends on the locale settings on which the code is running. When times and dates are displayed in the development environment, the short time format and short date format of the code locale are used. When displayed by running code, the short time format and short date format of the system locale are used, which may differ from the code locale. For this example, English/U.S. is assumed. `MyTime` and `MyDate` are displayed in the development environment using current system short time setting and short date setting.



```vb
Dim MyTime, MyDate, MyStr
MyTime = #17:04:23#
MyDate = #January 27, 1993#

' Returns current system time in the system-defined long time format.
MyStr = Format(Time, "Long Time")

' Returns current system date in the system-defined long date format.
MyStr = Format(Date, "Long Date")

MyStr = Format(MyTime, "h:m:s")    ' Returns "17:4:23".
MyStr = Format(MyTime, "hh:mm:ss AMPM")    ' Returns "05:04:23 PM".
MyStr = Format(MyDate, "dddd, mmm d yyyy")    ' Returns "Wednesday,
    ' Jan 27 1993".
' If format is not supplied, a string is returned.
MyStr = Format(23)    ' Returns "23".

' User-defined formats.
MyStr = Format(5459.4, "##,##0.00")    ' Returns "5,459.40".
MyStr = Format(334.9, "###0.00")    ' Returns "334.90".
MyStr = Format(5, "0.00%")    ' Returns "500.00%".
MyStr = Format("HELLO", "<")    ' Returns "hello".
MyStr = Format("This is it", ">")    ' Returns "THIS IS IT".
```

 **Different Formats for Different Numeric Values (Format Function)**
A user-defined format [expression](vbe-glossary.md) for numbers can have from one to four sections separated by semicolons. If the format[argument](vbe-glossary.md) contains one of the named numeric formats, only one section is allowed.


|**If you use**|**The result is**|
|:-----|:-----|
|One section only|The format expression applies to all values.|
|Two sections|The first section applies to positive values and zeros, the second to negative values.|
|Three sections|The first section applies to positive values, the second to negative values, and the third to zeros.|
|Four sections|The first section applies to positive values, the second to negative values, the third to zeros, and the fourth to [Null](vbe-glossary.md) values.|
The following example has two sections: the first defines the format for positive values and zeros; the second section defines the format for negative values.



```
"$#,##0;($#,##0)"
```

If you include semicolons with nothing between them, the missing section is printed using the format of the positive value. For example, the following format displays positive and negative values using the format in the first section and displays "Zero" if the value is zero.



```
"$#,##0;;\Z\e\r\o"
```

 **Different Formats for Different String Values (Format Function)**
A format [expression](vbe-glossary.md) for strings can have one section or two sections separated by a semicolon ( **;** ).


|**If you use**|**The result is**|
|:-----|:-----|
|One section only|The format applies to all string data.|
|Two sections|The first section applies to string data, the second to [Null](vbe-glossary.md) values and zero-length strings ("").|
 **Named Date/Time Formats (Format Function)**
The following table identifies the predefined date and time format names:


|**Format Name**|**Description**|
|:-----|:-----|
|**General Date**|Display a date and/or time. For real numbers, display a date and time, for example, 4/3/93 05:34 PM.If there is no fractional part, display only a date, for example, 4/3/93. If there is no integer part, display time only, for example, 05:34 PM. Date display is determined by your system settings.|
|**Long Date**|Display a date according to your system's long date format.|
|**Medium Date**|Display a date using the medium date format appropriate for the language version of the [host application](vbe-glossary.md).|
|**Short Date**|Display a date using your system's short date format.|
|**Long Time**|Display a time using your system's long time format; includes hours, minutes, seconds.|
|**Medium Time**|Display time in 12-hour format using hours and minutes and the AM/PM designator.|
|**Short Time**|Display a time using the 24-hour format, for example, 17:45.|
 **Named Numeric Formats (Format Function)**
The following table identifies the predefined numeric format names:


|**Format name**|**Description**|
|:-----|:-----|
|**General Number**|Display number with no thousand separator.|
|**Currency**|Display number with thousand separator, if appropriate; display two digits to the right of the decimal separator. Output is based on system [locale](vbe-glossary.md) settings.|
|**Fixed**|Display at least one digit to the left and two digits to the right of the decimal separator.|
|**Standard**|Display number with thousand separator, at least one digit to the left and two digits to the right of the decimal separator.|
|**Percent**|Display number multiplied by 100 with a percent sign ( **%** ) appended to the right; always display two digits to the right of the decimal separator.|
|**Scientific**|Use standard scientific notation.|
|**Yes/No**|Display No if number is 0; otherwise, display Yes.|
|**True/False**|Display  **False** if number is 0; otherwise, display **True**.|
|**On/Off**|Display Off if number is 0; otherwise, display On.|
 **User-Defined String Formats (Format Function)**
You can use any of the following characters to create a format [expression](vbe-glossary.md) for strings:


|**Character**|**Description**|
|:-----|:-----|
|**@**|Character placeholder. Display a character or a space. If the string has a character in the position where the at symbol ( **@** ) appears in the format string, display it; otherwise, display a space in that position. Placeholders are filled from right to left unless there is an exclamation point character ( **!** ) in the format string.|
|**&;**|Character placeholder. Display a character or nothing. If the string has a character in the position where the ampersand ( **&;** ) appears, display it; otherwise, display nothing. Placeholders are filled from right to left unless there is an exclamation point character ( **!** ) in the format string.|
|**&lt;**|Force lowercase. Display all characters in lowercase format.|
|**&gt;**|Force uppercase. Display all characters in uppercase format.|
|**!**|Force left to right fill of placeholders. The default is to fill placeholders from right to left.|
 **User-Defined Date/Time Formats (Format Function)**
The following table identifies characters you can use to create user-defined date/time formats:

|||
|:-----|:-----|
|**Character**|**Description**|
|( **:** )|Time separator. In some [locales](vbe-glossary.md), other characters may be used to represent the time separator. The time separator separates hours, minutes, and seconds when time values are formatted. The actual character used as the time separator in formatted output is determined by your system settings.|
|( **/** )|[Date separator](vbe-glossary.md). In some locales, other characters may be used to represent the date separator. The date separator separates the day, month, and year when date values are formatted. The actual character used as the date separator in formatted output is determined by your system settings.|
|c|Display the date as  `ddddd` and display the time as `ttttt`, in that order. Display only date information if there is no fractional part to the date serial number; display only time information if there is no integer portion.|
|d|Display the day as a number without a leading zero (1 - 31).|
|dd|Display the day as a number with a leading zero (01 - 31).|
|ddd|Display the day as an abbreviation (Sun - Sat).|
|dddd|Display the day as a full name (Sunday - Saturday).|
|ddddd|Display the date as a complete date (including day, month, and year), formatted according to your system's short date format setting. The default short date format is  `m/d/yy`.|
|dddddd|Display a date serial number as a complete date (including day, month, and year) formatted according to the long date setting recognized by your system. The default long date format is  `mmmm dd, yyyy`.|
|aaaa|The same as dddd, only it's the localized version of the string.|
|w|Display the day of the week as a number (1 for Sunday through 7 for Saturday).|
|ww|Display the week of the year as a number (1 - 54).|
|m|Display the month as a number without a leading zero (1 - 12). If  `m` immediately follows `h` or `hh`, the minute rather than the month is displayed.|
|mm|Display the month as a number with a leading zero (01 - 12). If  `m` immediately follows `h` or `hh`, the minute rather than the month is displayed. |
|mmm|Display the month as an abbreviation (Jan - Dec).|
|mmmm|Display the month as a full month name (January - December).|
|oooo|The same as mmmm, only it's the localized version of the string.|
|q|Display the quarter of the year as a number (1 - 4).|
|y|Display the day of the year as a number (1 - 366).|
|yy|Display the year as a 2-digit number (00 - 99).|
|yyyy|Display the year as a 4-digit number (100 - 9999).|
|h|Display the hour as a number without leading zeros (0 - 23).|
|Hh|Display the hour as a number with leading zeros (00 - 23).|
|N|Display the minute as a number without leading zeros (0 - 59).|
|Nn|Display the minute as a number with leading zeros (00 - 59).|
|S|Display the second as a number without leading zeros (0 - 59).|
|Ss|Display the second as a number with leading zeros (00 - 59).|
|t t t t t|Display a time as a complete time (including hour, minute, and second), formatted using the time separator defined by the time format recognized by your system. A leading zero is displayed if the leading zero option is selected and the time is before 10:00 A.M. or P.M. The default time format is  `h:mm:ss`.|
|AM/PM|Use the 12-hour clock and display an uppercase AM with any hour before noon; display an uppercase PM with any hour between noon and 11:59 P.M.|
|am/pm|Use the 12-hour clock and display a lowercase AM with any hour before noon; display a lowercase PM with any hour between noon and 11:59 P.M.|
|A/P|Use the 12-hour clock and display an uppercase A with any hour before noon; display an uppercase P with any hour between noon and 11:59 P.M.|
|a/p|Use the 12-hour clock and display a lowercase A with any hour before noon; display a lowercase P with any hour between noon and 11:59 P.M.|
|AMPM|Use the 12-hour clock and display the AM [string literal](vbe-glossary.md) as defined by your system with any hour before noon; display the PM string literal as defined by your system with any hour between noon and 11:59 P.M. AMPM can be either uppercase or lowercase, but the case of the string displayed matches the string as defined by your system settings. The default format is AM/PM.|
 **User-Defined Numeric Formats (Format Function)**
The following table identifies characters you can use to create user-defined number formats:

|||
|:-----|:-----|
|Character|Description|
|None|Display the number with no formatting.|
|( **0** )|Digit placeholder. Display a digit or a zero. If the [expression](vbe-glossary.md) has a digit in the position where the 0 appears in the format string, display it; otherwise, display a zero in that position.If the number has fewer digits than there are zeros (on either side of the decimal) in the format expression, display leading or trailing zeros. If the number has more digits to the right of the decimal separator than there are zeros to the right of the decimal separator in the format expression, round the number to as many decimal places as there are zeros. If the number has more digits to the left of the decimal separator than there are zeros to the left of the decimal separator in the format expression, display the extra digits without modification.|
|( **#** )|Digit placeholder. Display a digit or nothing. If the expression has a digit in the position where the # appears in the format string, display it; otherwise, display nothing in that position. This symbol works like the 0 digit placeholder, except that leading and trailing zeros aren't displayed if the number has the same or fewer digits than there are # characters on either side of the decimal separator in the format expression.|
|( **.** )|Decimal placeholder. In some [locales](vbe-glossary.md), a comma is used as the decimal separator. The decimal placeholder determines how many digits are displayed to the left and right of the decimal separator. If the format expression contains only number signs to the left of this symbol, numbers smaller than 1 begin with a decimal separator. To display a leading zero displayed with fractional numbers, use 0 as the first digit placeholder to the left of the decimal separator. The actual character used as a decimal placeholder in the formatted output depends on the Number Format recognized by your system.|
|( **%)**|Percentage placeholder. The expression is multiplied by 100. The percent character ( **%** ) is inserted in the position where it appears in the format string.|
|( **,** )|Thousand separator. In some locales, a period is used as a thousand separator. The thousand separator separates thousands from hundreds within a number that has four or more places to the left of the decimal separator. Standard use of the thousand separator is specified if the format contains a thousand separator surrounded by digit placeholders ( **0** or **#** ). Two adjacent thousand separators or a thousand separator immediately to the left of the decimal separator (whether or not a decimal is specified) means "scale the number by dividing it by 1000, rounding as needed." For example, you can use the format string "##0,," to represent 100 million as 100. Numbers smaller than 1 million are displayed as 0. Two adjacent thousand separators in any position other than immediately to the left of the decimal separator are treated simply as specifying the use of a thousand separator. The actual character used as the thousand separator in the formatted output depends on the Number Format recognized by your system.|
|( **:** )|Time separator. In some locales, other characters may be used to represent the time separator. The time separator separates hours, minutes, and seconds when time values are formatted. The actual character used as the time separator in formatted output is determined by your system settings.|
|( **/** )|[Date separator](vbe-glossary.md). In some locales, other characters may be used to represent the date separator. The date separator separates the day, month, and year when date values are formatted. The actual character used as the date separator in formatted output is determined by your system settings.|
|( **E- E+ e- e+** )|Scientific format. If the format expression contains at least one digit placeholder ( **0** or **#** ) to the right of E-, E+, e-, or e+, the number is displayed in scientific format and E or e is inserted between the number and its exponent. The number of digit placeholders to the right determines the number of digits in the exponent. Use E- or e- to place a minus sign next to negative exponents. Use E+ or e+ to place a minus sign next to negative exponents and a plus sign next to positive exponents.|
|**- + $** ( )|Display a literal character. To display a character other than one of those listed, precede it with a backslash (\) or enclose it in double quotation marks (" ").|
|( **\** )|Display the next character in the format string. To display a character that has special meaning as a literal character, precede it with a backslash (\). The backslash itself isn't displayed. Using a backslash is the same as enclosing the next character in double quotation marks. To display a backslash, use two backslashes (\\). Examples of characters that can't be displayed as literal characters are the date-formatting and time-formatting characters (a, c, d, h, m, n, p, q, s, t, w, y, / and :), the numeric-formatting characters (#, 0, %, E, e, comma, and period), and the string-formatting characters (@, &;, <, >, and !).|
|("ABC")|Display the string inside the double quotation marks (" "). To include a string in  **_format_** from within code, you must use **Chr(** 34 **)** to enclose the text (34 is the[character code](vbe-glossary.md) for a quotation mark (")).|

