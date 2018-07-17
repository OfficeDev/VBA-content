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


|**Part**|**Description**|
|:-----|:-----|
|**_expression_**|Required. [String expression](vbe-glossary.md) containing substring to replace.|
|**_find_**|Required. Substring being searched for.|
|**_replace_**|Required. Replacement substring.|
|**_start_**|Optional. Position within  **_expression_** where substring search is to begin. If omitted, 1 is assumed.|
|**_count_**|Optional. Number of substring substitutions to perform. If omitted, the default value is -1, which means make all possible substitutions.|
|**_compare_**|Optional. Numeric value indicating the kind of comparison to use when evaluating substrings. See Settings section for values.|
 **Settings**
The  **_compare_** argument can have the following values:


|**Constant**|**Value**|**Description**|
|:-----|:-----|:-----|
|**vbUseCompareOption**|-1|Performs a comparison using the setting of the  **Option Compare** statement.|
|**vbBinaryCompare**|0|Performs a binary comparison.|
|**vbTextCompare**|1|Performs a textual comparison.|
|**vbDatabaseCompare**|2|Microsoft Access only. Performs a comparison based on information in your database.|
 **Return Values**
 **Replace** returns the following values:


|**If**|**Replace returns**|
|:-----|:-----|
|**_expression_** is zero-length|Zero-length string ("")|
|**_expression_** is **Null**|An error.|
|**_find_** is zero-length|Copy of  **_expression_**.|
|**_replace_** is zero-length|Copy of  **_expression_** with all occurences of _find_ removed.|
|**_start_** > **Len( _expression_ )**|Zero-length string.|
|**_count_** is 0|Copy of  **_expression_**.|
 **Remarks**
The return value of the  **Replace** function is a string, with substitutions made, that begins at the position specified by **_start_** and and concludes at the end of the **_expression_** string. It is not a copy of the original string from start to finish.

