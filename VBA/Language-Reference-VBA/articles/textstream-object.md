---
title: TextStream Object
keywords: vblr6.chm2181930
f1_keywords:
- vblr6.chm2181930
ms.prod: office
api_name:
- Office.TextStream
ms.assetid: b1b78d3a-78b3-aee5-2efc-1e208e0858ac
ms.date: 06/08/2017
---


# TextStream Object



 **Description**
Facilitates sequential access to file.
 **Syntax**
 **TextStream.** { _property_ | _method_ }
The  _property_ and _method_ arguments can be any of the properties and methods associated with the **TextStream** object. Note that in actual usage **TextStream** is replaced by a variable placeholder representing the **TextStream** object returned from the **FileSystemObject**.
 **Remarks**
In the following code,  `a` is the **TextStream** object returned by the **CreateTextFile** method on the **FileSystemObject**:
 **WriteLine** and **Close** are two methods of the **TextStream** Object.



```vb
Set fs = CreateObject("Scripting.FileSystemObject")
Set a = fs.CreateTextFile("c:\testfile.txt", True)
a.WriteLine("This is a test.")
a.Close

```


