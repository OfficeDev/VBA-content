---
title: FileLen Function
keywords: vblr6.chm1008922
f1_keywords:
- vblr6.chm1008922
ms.prod: office
ms.assetid: 019f4538-9d04-d8f9-4689-0e36ac32a753
ms.date: 06/08/2017
---


# FileLen Function



Returns a [Long](vbe-glossary.md) specifying the length of a file in bytes.
 **Syntax**
 **FileLen(**_pathname_**)**
The required  _pathname_[argument](vbe-glossary.md) is a[string expression](vbe-glossary.md) that specifies a file. The _pathname_ may include the directory or folder, and the drive.
 **Remarks**
If the specified file is open when the  **FileLen** function is called, the value returned represents the size of the file immediately before it was opened.

 **Note**  To obtain the length of an open file, use the  **LOF** function.


## Example

This example uses the  **FileLen** function to return the length of a file in bytes. For purposes of this example, assume that `TESTFILE` is a file containing some data.


```vb
Dim MySize
MySize = FileLen("TESTFILE")    ' Returns file length (bytes).


```


