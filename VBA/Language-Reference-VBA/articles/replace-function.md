---
title: Replace Function
keywords: vblr6.chm1008930
f1_keywords:
- vblr6.chm1008930
ms.prod: office
ms.assetid: a24e3da4-fc94-56e7-d718-f4c2d0a31072
ms.date: 06/08/2017
---


# Replace Function



 **Description**
Returns a string in which a specified substring has been replaced with another substring a specified number of times.
 **Syntax**
 **Replace( _expression_,** **_find_,** **_replace_** [ **,** **_start_** [ **,** **_count_** [ **,** **_compare_** ]]] **)**
The  **Replace** function syntax has these[named arguments](vbe-glossary.md):


| <strong>Part</strong>                | <strong>Description</strong>                                                                                                              |
|:-------------------------------------|:------------------------------------------------------------------------------------------------------------------------------------------|
| <strong><em>expression</em></strong> | Required. [String expression](vbe-glossary.md) containing substring to replace.                                                           |
| <strong><em>find</em></strong>       | Required. Substring being searched for.                                                                                                   |
| <strong><em>replace</em></strong>    | Required. Replacement substring.                                                                                                          |
| <strong><em>start</em></strong>      | Optional. Position within  <strong><em>expression</em></strong> where substring search is to begin. If omitted, 1 is assumed.             |
| <strong><em>count</em></strong>      | Optional. Number of substring substitutions to perform. If omitted, the default value is -1, which means make all possible substitutions. |
| <strong><em>compare</em></strong>    | Optional. Numeric value indicating the kind of comparison to use when evaluating substrings. See Settings section for values.             |

 **Settings**
The  **_compare_** argument can have the following values:


| <strong>Constant</strong>           | <strong>Value</strong> | <strong>Description</strong>                                                               |
|:------------------------------------|:-----------------------|:-------------------------------------------------------------------------------------------|
| <strong>vbUseCompareOption</strong> | -1                     | Performs a comparison using the setting of the  <strong>Option Compare</strong> statement. |
| <strong>vbBinaryCompare</strong>    | 0                      | Performs a binary comparison.                                                              |
| <strong>vbTextCompare</strong>      | 1                      | Performs a textual comparison.                                                             |
| <strong>vbDatabaseCompare</strong>  | 2                      | Microsoft Access only. Performs a comparison based on information in your database.        |

 **Return Values**
 **Replace** returns the following values:


| <strong>If</strong>                                                           | <strong>Replace returns</strong>                                                            |
|:------------------------------------------------------------------------------|:--------------------------------------------------------------------------------------------|
| <strong><em>expression</em></strong> is zero-length                           | Zero-length string ("")                                                                     |
| <strong><em>expression</em></strong> is <strong>Null</strong>                 | An error.                                                                                   |
| <strong><em>find</em></strong> is zero-length                                 | Copy of  <strong><em>expression</em></strong>.                                              |
| <strong><em>replace</em></strong> is zero-length                              | Copy of  <strong><em>expression</em></strong> with all occurences of <em>find</em> removed. |
| <strong><em>start</em></strong> > <strong>Len( <em>expression</em> )</strong> | Zero-length string.                                                                         |
| <strong><em>count</em></strong> is 0                                          | Copy of  <strong><em>expression</em></strong>.                                              |

 **Remarks**
The return value of the  **Replace** function is a string, with substitutions made, that begins at the position specified by **_start_** and and concludes at the end of the **_expression_** string. It is not a copy of the original string from start to finish.

