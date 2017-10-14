---
title: InStr Function
keywords: vblr6.chm1008946
f1_keywords:
- vblr6.chm1008946
ms.prod: office
ms.assetid: d83b314a-e77c-fc18-0744-266f982a82b7
ms.date: 06/08/2017
---


# InStr Function



Returns a  **Variant** ( **Long** ) specifying the position of the first occurrence of one string within another.
 **Syntax**
 **InStr** ([ _start_, ] _string1_, _string2_ [, _compare_ ])
The  **InStr** function syntax has these[arguments](vbe-glossary.md):


|**Part**|**Description**|
|:-----|:-----|
| _start_|Optional. [Numeric expression](vbe-glossary.md) that sets the starting position for each search. If omitted, search begins at the first character position. If **_start_** contains[Null](vbe-glossary.md), an error occurs. The  **_start_** argument is required if **_compare_** is specified.|
| _string1_|Required. [String expression](vbe-glossary.md) being searched.|
| _string2_|Required. String expression sought.|
| _compare_|Optional. Specifies the type of [string comparison](vbe-glossary.md). If  **_compare_** is Null, an error occurs. If **_compare_** is omitted, the **Option** **Compare** setting determines the type of comparison. Specify a valid LCID (LocaleID) to use locale-specific rules in the comparison.|
 **Settings**
The  _compare_ argument settings are:


|**Constant**|**Value**|**Description**|
|:-----|:-----|:-----|
|**vbUseCompareOption**|-1|Performs a comparison using the setting of the  **Option Compare** statement.|
|**vbBinaryCompare**|0|Performs a binary comparison.|
|**vbTextCompare**|1|Performs a textual comparison.|
|**vbDatabaseCompare**|2|Microsoft Access only. Performs a comparison based on information in your database.|
 **Return Values**


|**If**|**InStr returns**|
|:-----|:-----|
|**_string1_** is zero-length|0|
|**_string1_** is **Null**|Null|
|**_string2_** is zero-length|**_start_**|
|**_string2_** is **Null**|Null|
|**_string2_** is not found|0|
|**_string2_** is found within **_string1_**|Position at which match is found|
|**_start_** > **_string2_**|0|
 **Remarks**
The  **InStrB** function is used with byte data contained in a string. Instead of returning the character position of the first occurrence of one string within another, **InStrB** returns the byte position.

## Example

This example uses the  **InStr** function to return the position of the first occurrence of one string within another.


```vb
Dim SearchString, SearchChar, MyPos
SearchString ="XXpXXpXXPXXP"    ' String to search in.
SearchChar = "P"    ' Search for "P".

' A textual comparison starting at position 4. Returns 6.
MyPos = Instr(4, SearchString, SearchChar, 1)    

' A binary comparison starting at position 1. Returns 9.
MyPos = Instr(1, SearchString, SearchChar, 0)

' Comparison is binary by default (last argument is omitted).
MyPos = Instr(SearchString, SearchChar)    ' Returns 9.

MyPos = Instr(1, SearchString, "W")    ' Returns 0.
```


