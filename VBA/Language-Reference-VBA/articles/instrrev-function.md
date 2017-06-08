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


|**Part**|**Description**|
|:-----|:-----|
|**_stringcheck_**|Required. [String expression](vbe-glossary.md) being searched.|
|**_stringmatch_**|Required. String expression being searched for.|
|**_start_**|Optional. [Numeric expression](vbe-glossary.md) that sets the starting position for each search. If omitted, -1 is used, which means that the search begins at the last character position. If **_start_** contains[Null](vbe-glossary.md), an error occurs.|
|**_compare_**|Optional. Numeric value indicating the kind of comparison to use when evaluating substrings. If omitted, a binary comparison is performed. See Settings section for values.|
 **Settings**
The  **_compare_** argument can have the following values:


|**Constant**|**Value**|**Description**|
|:-----|:-----|:-----|
|**vbUseCompareOption**|-1|Performs a comparison using the setting of the  **Option Compare** statement.|
|**vbBinaryCompare**| 0|Performs a binary comparison.|
|**vbTextCompare**| 1|Performs a textual comparison.|
|**vbDatabaseCompare**| 2|Microsoft Access only. Performs a comparison based on information in your database.|
 **Return Values**
 **InStrRev** returns the following values:


|**If**|**InStrRev returns**|
|:-----|:-----|
|**_stringcheck_** is zero-length|0|
|**_stringcheck_** is **Null**|**Null**|
|**_stringmatch_** is zero-length| _start_|
|**_stringmatch_** is **Null**|**Null**|
|**_stringmatch_** is not found|0|
|**_stringmatch_** is found within **_stringcheck_**|Position at which match is found|
|**_start_** > **Len( _stringmatch_ )**|0|
 **Remarks**
Note that the syntax for the  **InstrRev** function is not the same as the syntax for the **Instr** function.

