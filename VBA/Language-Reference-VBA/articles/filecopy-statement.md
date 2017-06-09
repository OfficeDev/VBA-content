---
title: FileCopy Statement
keywords: vblr6.chm1008920
f1_keywords:
- vblr6.chm1008920
ms.prod: office
ms.assetid: 9da94e6e-f8c4-70cd-40b5-501668cbfd71
ms.date: 06/08/2017
---


# FileCopy Statement

Copies a file.

 **Syntax**

 **FileCopy** **_source_,** **_destination_**

The  **FileCopy** statement syntax has these[named arguments](vbe-glossary.md):


|**Part**|**Description**|
|:-----|:-----|
|**_source_**|Required. [String expression](vbe-glossary.md) that specifies the name of the file to be copied. The **_source_** may include directory or folder, and drive.|
|**_destination_**|Required. String expression that specifies the target file name. The  **_destination_** may include directory or folder, and drive.|
 **Remarks**
If you try to use the  **FileCopy** statement on a currently open file, an error occurs.

## Example

This example uses the  **FileCopy** statement to copy one file to another. For purposes of this example, assume that is a file containing some data.


```vb
Dim SourceFile, DestinationFile 
SourceFile = "SRCFILE" ' Define source file name. 
DestinationFile = "DESTFILE" ' Define target file name. 
FileCopy SourceFile, DestinationFile ' Copy source to target. 

```


