---
title: InStrRev Function
keywords: vblr6.chm1008911
f1_keywords:
- vblr6.chm1008911
ms.prod: office
ms.assetid: 2677e5dc-a128-1bf4-dd72-304469b46cc2
ms.date: 06/08/2017
---


# InStrRev Function



 **Description**
Returns the position of an occurrence of one string within another, from the end of string.
 **Syntax**
 **InstrRev( _stringcheck_,** **_stringmatch_** [ **,** **_start_** [ **,** **_compare_** ]] **)**
The  **InstrRev** function syntax has these[named arguments](vbe-glossary.md):


| <strong>Part</strong>                 | <strong>Description</strong>                                                                                                                                                                                                                                                     |
|:--------------------------------------|:---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------|
| <strong><em>stringcheck</em></strong> | Required. [String expression](vbe-glossary.md) being searched.                                                                                                                                                                                                                   |
| <strong><em>stringmatch</em></strong> | Required. String expression being searched for.                                                                                                                                                                                                                                  |
| <strong><em>start</em></strong>       | Optional. [Numeric expression](vbe-glossary.md) that sets the starting position for each search. If omitted, -1 is used, which means that the search begins at the last character position. If <strong><em>start</em></strong> contains[Null](vbe-glossary.md), an error occurs. |
| <strong><em>compare</em></strong>     | Optional. Numeric value indicating the kind of comparison to use when evaluating substrings. If omitted, a binary comparison is performed. See Settings section for values.                                                                                                      |

 **Settings**
The  **_compare_** argument can have the following values:


| <strong>Constant</strong>           | <strong>Value</strong> | <strong>Description</strong>                                                               |
|:------------------------------------|:-----------------------|:-------------------------------------------------------------------------------------------|
| <strong>vbUseCompareOption</strong> | -1                     | Performs a comparison using the setting of the  <strong>Option Compare</strong> statement. |
| <strong>vbBinaryCompare</strong>    | 0                      | Performs a binary comparison.                                                              |
| <strong>vbTextCompare</strong>      | 1                      | Performs a textual comparison.                                                             |
| <strong>vbDatabaseCompare</strong>  | 2                      | Microsoft Access only. Performs a comparison based on information in your database.        |

 **Return Values**
 **InStrRev** returns the following values:


| <strong>If</strong>                                                                         | <strong>InStrRev returns</strong> |
|:--------------------------------------------------------------------------------------------|:----------------------------------|
| <strong><em>stringcheck</em></strong> is zero-length                                        | 0                                 |
| <strong><em>stringcheck</em></strong> is <strong>Null</strong>                              | <strong>Null</strong>             |
| <strong><em>stringmatch</em></strong> is zero-length                                        | <em>start</em>                    |
| <strong><em>stringmatch</em></strong> is <strong>Null</strong>                              | <strong>Null</strong>             |
| <strong><em>stringmatch</em></strong> is not found                                          | 0                                 |
| <strong><em>stringmatch</em></strong> is found within <strong><em>stringcheck</em></strong> | Position at which match is found  |
| <strong><em>start</em></strong> > <strong>Len( <em>stringmatch</em> )</strong>              | 0                                 |

 **Remarks**
Note that the syntax for the  **InstrRev** function is not the same as the syntax for the **Instr** function.

