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


| <strong>Part</strong> | <strong>Description</strong>                                                                                                                                                                                                                                                                                                                                           |
|:----------------------|:-----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------|
| <em>start</em>        | Optional. [Numeric expression](vbe-glossary.md) that sets the starting position for each search. If omitted, search begins at the first character position. If <strong><em>start</em></strong> contains[Null](vbe-glossary.md), an error occurs. The  <strong><em>start</em></strong> argument is required if <strong><em>compare</em></strong> is specified.          |
| <em>string1</em>      | Required. [String expression](vbe-glossary.md) being searched.                                                                                                                                                                                                                                                                                                         |
| <em>string2</em>      | Required. String expression sought.                                                                                                                                                                                                                                                                                                                                    |
| <em>compare</em>      | Optional. Specifies the type of [string comparison](vbe-glossary.md). If  <strong><em>compare</em></strong> is Null, an error occurs. If <strong><em>compare</em></strong> is omitted, the <strong>Option</strong> <strong>Compare</strong> setting determines the type of comparison. Specify a valid LCID (LocaleID) to use locale-specific rules in the comparison. |

 **Settings**
The  _compare_ argument settings are:


| <strong>Constant</strong>           | <strong>Value</strong> | <strong>Description</strong>                                                               |
|:------------------------------------|:-----------------------|:-------------------------------------------------------------------------------------------|
| <strong>vbUseCompareOption</strong> | -1                     | Performs a comparison using the setting of the  <strong>Option Compare</strong> statement. |
| <strong>vbBinaryCompare</strong>    | 0                      | Performs a binary comparison.                                                              |
| <strong>vbTextCompare</strong>      | 1                      | Performs a textual comparison.                                                             |
| <strong>vbDatabaseCompare</strong>  | 2                      | Microsoft Access only. Performs a comparison based on information in your database.        |

 **Return Values**


| <strong>If</strong>                                                                 | <strong>InStr returns</strong>   |
|:------------------------------------------------------------------------------------|:---------------------------------|
| <strong><em>string1</em></strong> is zero-length                                    | 0                                |
| <strong><em>string1</em></strong> is <strong>Null</strong>                          | Null                             |
| <strong><em>string2</em></strong> is zero-length                                    | <strong><em>start</em></strong>  |
| <strong><em>string2</em></strong> is <strong>Null</strong>                          | Null                             |
| <strong><em>string2</em></strong> is not found                                      | 0                                |
| <strong><em>string2</em></strong> is found within <strong><em>string1</em></strong> | Position at which match is found |
| <strong><em>start</em></strong> > <strong><em>string2</em></strong>                 | 0                                |

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


