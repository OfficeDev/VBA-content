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


|**Part**|**Description**|
|:-----|:-----|
|**_sourcearray_**|Required. One-dimensional array of strings to be searched.|
|**_match_**|Required. String to search for.|
|**_include_**|Optional.  **Boolean** value indicating whether to return substrings that include or exclude **_match_**. If **_include_** is **True**, **Filter** returns the subset of the array that contains **_match_** as a substring. If **_include_** is **False**, **Filter** returns the subset of the array that does not contain **_match_** as a substring.|
|**_compare_**|Optional. Numeric value indicating the kind of string comparison to use. See Settings section for values.|
 **Settings**
The  **_compare_** argument can have the following values:


|**Constant**|**Value**|**Description**|
|:-----|:-----|:-----|
|**vbUseCompareOption**|-1|Performs a comparison using the setting of the  **Option Compare** statement.|
|**vbBinaryCompare**| 0|Performs a binary comparison.|
|**vbTextCompare**| 1|Performs a textual comparison.|
|**vbDatabaseCompare**| 2|Microsoft Access only. Performs a comparison based on information in your database.|
 **Remarks**
If no matches of  **_match_** are found within **_sourcearray_**, **Filter** returns an empty array. An error occurs if **_sourcearray_** is **Null** or is not a one-dimensional array.
The array returned by the  **Filter** function contains only enough elements to contain the number of matched items.

