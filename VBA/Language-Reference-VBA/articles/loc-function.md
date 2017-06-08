---
title: Loc Function
keywords: vblr6.chm1008963
f1_keywords:
- vblr6.chm1008963
ms.prod: office
ms.assetid: e744813a-3633-e6d1-4f4c-517f1dcec196
ms.date: 06/08/2017
---


# Loc Function



Returns a [Long](vbe-glossary.md) specifying the current read/write position within an open file.
 **Syntax**
 **Loc(**_filenumber_**)**
The required  _filenumber_[argument](vbe-glossary.md) is any valid[Integer](vbe-glossary.md)[file number](vbe-glossary.md).
 **Remarks**
The following describes the return value for each file access mode:


|**Mode**|**Return Value**|
|:-----|:-----|
|**Random**|Number of the last record read from or written to the file.|
|**Sequential**|Current byte position in the file divided by 128. However, information returned by  **Loc** for sequential files is neither used nor required.|
|**Binary**|Position of the last byte read or written.|

## Example

This example uses the  **Loc** function to return the current read/write position within an open file. This example assumes that `TESTFILE` is a text file with a few lines of sample data.


```vb
Dim MyLocation, MyLine
Open "TESTFILE" For Binary As #1    ' Open file just created.
Do While MyLocation < LOF(1)    ' Loop until end of file.
    MyLine = MyLine &; Input(1, #1)    ' Read character into variable.
    MyLocation = Loc(1)    ' Get current position within file.
' Print to the Immediate window.
    Debug.Print MyLine; Tab; MyLocation
Loop
Close #1    ' Close file.

```


