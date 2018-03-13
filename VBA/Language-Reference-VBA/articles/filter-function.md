---
title: Filter Function
keywords: vblr6.chm1008912
f1_keywords:
- vblr6.chm1008912
ms.prod: office
ms.assetid: 00630b25-e7b8-5c32-b6d1-9816f01c3a0f
ms.date: 06/08/2017
---


# Filter Function



 **Description**
Returns a zero-based array containing subset of a string array based on a specified filter criteria.
 **Syntax**
 **Filter( _sourcearray, match_** [ **_, include_** [ **_, compare_** ]] **)**
The  **Filter** function syntax has these[named argument](vbe-glossary.md):


| <strong>Part</strong>                 | <strong>Description</strong>                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                  |
|:--------------------------------------|:------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------|
| <strong><em>sourcearray</em></strong> | Required. One-dimensional array of strings to be searched.                                                                                                                                                                                                                                                                                                                                                                                                                                                                    |
| <strong><em>match</em></strong>       | Required. String to search for.                                                                                                                                                                                                                                                                                                                                                                                                                                                                                               |
| <strong><em>include</em></strong>     | Optional.  <strong>Boolean</strong> value indicating whether to return substrings that include or exclude <strong><em>match</em></strong>. If <strong><em>include</em></strong> is <strong>True</strong>, <strong>Filter</strong> returns the subset of the array that contains <strong><em>match</em></strong> as a substring. If <strong><em>include</em></strong> is <strong>False</strong>, <strong>Filter</strong> returns the subset of the array that does not contain <strong><em>match</em></strong> as a substring. |
| <strong><em>compare</em></strong>     | Optional. Numeric value indicating the kind of string comparison to use. See Settings section for values.                                                                                                                                                                                                                                                                                                                                                                                                                     |

 **Settings**
The  **_compare_** argument can have the following values:


| <strong>Constant</strong>           | <strong>Value</strong> | <strong>Description</strong>                                                               |
|:------------------------------------|:-----------------------|:-------------------------------------------------------------------------------------------|
| <strong>vbUseCompareOption</strong> | -1                     | Performs a comparison using the setting of the  <strong>Option Compare</strong> statement. |
| <strong>vbBinaryCompare</strong>    | 0                      | Performs a binary comparison.                                                              |
| <strong>vbTextCompare</strong>      | 1                      | Performs a textual comparison.                                                             |
| <strong>vbDatabaseCompare</strong>  | 2                      | Microsoft Access only. Performs a comparison based on information in your database.        |

 **Remarks**
If no matches of  **_match_** are found within **_sourcearray_**, **Filter** returns an empty array. An error occurs if **_sourcearray_** is **Null** or is not a one-dimensional array.
The array returned by the  **Filter** function contains only enough elements to contain the number of matched items.

