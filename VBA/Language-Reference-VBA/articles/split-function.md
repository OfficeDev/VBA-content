---
title: Split Function
keywords: vblr6.chm1008907
f1_keywords:
- vblr6.chm1008907
ms.prod: office
ms.assetid: 7c68f50a-c4c4-ee16-cc04-9d067a0b5819
ms.date: 06/08/2017
---


# Split Function



 **Description**
Returns a zero-based, one-dimensional [array](vbe-glossary.md) containing a specified number of substrings.
 **Syntax**
 **Split( _expression_** [ **,** **_delimiter_** [ **,** **_limit_** [ **,** **_compare_** ]]] **)**
The  **Split** function syntax has these[named arguments](vbe-glossary.md):


|**Part**|**Description**|
|:-----|:-----|
|**_expression_**|Required. [String expression](vbe-glossary.md) containing substrings and delimiters. If _expression_ is a zero-length string(""), **Split** returns an empty array, that is, an array with no elements and no data.|
|**_delimiter_**|Optional. String character used to identify substring limits. If omitted, the space character (" ") is assumed to be the delimiter. If  **_delimiter_** is a zero-length string, a single-element array containing the entire **_expression_** string is returned.|
|**_limit_**|Optional. Number of substrings to be returned; -1 indicates that all substrings are returned.|
|**_compare_**|Optional. Numeric value indicating the kind of comparison to use when evaluating substrings. See Settings section for values.|
 **Settings**
The  **_compare_** argument can have the following values:


|**Constant**|**Value**|**Description**|
|:-----|:-----|:-----|
|**vbUseCompareOption**|-1|Performs a comparison using the setting of the  **Option Compare** statement.|
|**vbBinaryCompare**|0|Performs a binary comparison.|
|**vbTextCompare**|1|Performs a textual comparison.|
|**vbDatabaseCompare**|2|Microsoft Access only. Performs a comparison based on information in your database.|

