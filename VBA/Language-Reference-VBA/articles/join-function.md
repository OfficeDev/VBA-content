---
title: Join Function
keywords: vblr6.chm1008915
f1_keywords:
- vblr6.chm1008915
ms.prod: office
ms.assetid: 2c7a6ee5-ea52-1f93-1f16-20e333804b23
ms.date: 06/08/2017
---


# Join Function



 **Description**
Returns a string created by joining a number of substrings contained in an [array](vbe-glossary.md).
 **Syntax**
 **Join( _sourcearray_** [, **_delimiter_** ] **)**
The  **Join** function syntax has these[named arguments](vbe-glossary.md):


|**Part**|**Description**|
|:-----|:-----|
|**_sourcearray_**|Required. One-dimensional array containing substrings to be joined.|
|**_delimiter_**|Optional. String character used to separate the substrings in the returned string. If omitted, the space character (" ") is used. If  **_delimiter_** is a zero-length string (""), all items in the list are concatenated with no delimiters.|

