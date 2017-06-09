---
title: AtEndOfStream Property
keywords: vblr6.chm2182072
f1_keywords:
- vblr6.chm2182072
ms.prod: office
api_name:
- Office.AtEndOfStream
ms.assetid: 157b18dc-fdfb-a9f6-1368-aaf4654a2ef5
ms.date: 06/08/2017
---


# AtEndOfStream Property



 **Description**
Read-only property that returns  **True** if the file pointer is at the end of a **TextStream** file; **False** if it is not.
 **Syntax**
 _object_. **AtEndOfStream**
The  _object_ is always the name of a **TextStream** object.
 **Remarks**
The  **AtEndOfStream** property applies only to **TextStream** files that are open for reading; otherwise, an error occurs.
The following code illustrates the use of the  **AtEndOfStream** property:



```vb
Dim fs, a, retstring
Set fs = CreateObject("Scripting.FileSystemObject")
Set a = fs.OpenTextFile("c:\testfile.txt", ForReading, False)
Do While a. AtEndOfStream <> True
    retstring = a.ReadLine
    ...
Loop
a.Close

```


