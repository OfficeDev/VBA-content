---
title: LOF Function
keywords: vblr6.chm1008965
f1_keywords:
- vblr6.chm1008965
ms.prod: office
ms.assetid: 1bf66bce-d3d7-9c34-e8d2-8ad1e1ee24a8
ms.date: 06/08/2017
---


# LOF Function



Returns a [Long](vbe-glossary.md) representing the size, in bytes, of a file opened using the **Open** statement.
 **Syntax**
 **LOF(**_filenumber_**)**
The required  _filenumber_[argument](vbe-glossary.md) is an[Integer](vbe-glossary.md) containing a valid[file number](vbe-glossary.md).

 **Note**  Use the  **FileLen** function to obtain the length of a file that is not open.


## Example

This example uses the  **LOF** function to determine the size of an open file. This example assumes that `TESTFILE` is a text file containing sample data.


```vb
Dim FileLength
Open "TESTFILE" For Input As #1    ' Open file.
FileLength = LOF(1)    ' Get length of file.
Close #1    ' Close file.


```


