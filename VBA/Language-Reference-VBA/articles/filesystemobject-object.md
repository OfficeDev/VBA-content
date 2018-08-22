---
title: FileSystemObject Object
keywords: vblr6.chm2181927
f1_keywords:
- vblr6.chm2181927
ms.prod: office
api_name:
- Office.FileSystemObject
ms.assetid: 7ad2dad3-c6d8-90a6-77a5-c712da8316f3
ms.date: 06/08/2017
---


# FileSystemObject Object

Provides access to a computer's file system.

## Syntax

 **Scripting.FileSystemObject**

## Remarks

The following code illustrates how the  **FileSystemObject** is used to return a **[TextStream](textstream-object.md)** object that can be read from or written to:



```vb
Set fs = CreateObject("Scripting.FileSystemObject")
Set a = fs.CreateTextFile("c:\testfile.txt", True)
a.WriteLine("This is a test.")
a.Close
```

In the code shown above, the  **[CreateObject](createobject-function.md)** function returns the **FileSystemObject** ( `fs` ). The **[CreateTextFile](createtextfile-method.md)** method then creates the file as a **[TextStream](textstream-object.md)** object ( `a` ), and the **[WriteLine](writeline-method.md)** method writes a line of text to the created text file. The **[Close](close-method-filesystemobject-object.md)** method flushes the buffer and closes the file.

