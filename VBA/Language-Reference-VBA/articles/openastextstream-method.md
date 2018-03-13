---
title: OpenAsTextStream Method
keywords: vblr6.chm2182007
f1_keywords:
- vblr6.chm2182007
ms.prod: office
api_name:
- Office.OpenAsTextStream
ms.assetid: 11bdf601-368b-7d95-a7db-394271d59da6
ms.date: 06/08/2017
---


# OpenAsTextStream Method



 **Description**
Opens a specified file and returns a  **TextStream** object that can be used to read from, write to, or append to the file.
 **Syntax**
 _object_. **OpenAsTextStream(** [ _iomode_, [ _format_ ]] **)**
The  **OpenAsTextStream** method syntax has these parts:


| <strong>Part</strong> | <strong>Description</strong>                                                                                                                                       |
|:----------------------|:-------------------------------------------------------------------------------------------------------------------------------------------------------------------|
| <em>object</em>       | Required. Always the name of a  <strong>File</strong> object.                                                                                                      |
| <em>iomode</em>       | Optional. Indicates input/output mode. Can be one of three constants:  <strong>ForReading</strong>, <strong>ForWriting</strong>, or <strong>ForAppending</strong>. |
| <em>format</em>       | Optional. One of three  <strong>Tristate</strong> values used to indicate the format of the opened file. If omitted, the file is opened as ASCII.                  |

 **Settings**
The  _iomode_ argument can have any of the following settings:


| <strong>Constant</strong>     | <strong>Value</strong> | <strong>Description</strong>                                                                         |
|:------------------------------|:-----------------------|:-----------------------------------------------------------------------------------------------------|
| <strong>ForReading</strong>   | 1                      | Open a file for reading only. You can't write to this file.                                          |
| <strong>ForWriting</strong>   | 2                      | Open a file for writing. If a file with the same name exists, its previous contents are overwritten. |
| <strong>ForAppending</strong> | 8                      | Open a file and write to the end of the file.                                                        |

The  _format_ argument can have any of the following settings:


| <strong>Constant</strong>           | <strong>Value</strong> | <strong>Description</strong>             |
|:------------------------------------|:-----------------------|:-----------------------------------------|
| <strong>TristateUseDefault</strong> | -2                     | Opens the file using the system default. |
| <strong>TristateTrue</strong>       | -1                     | Opens the file as Unicode.               |
| <strong>TristateFalse</strong>      | 0                      | Opens the file as ASCII.                 |

 **Remarks**
The  **OpenAsTextStream** method provides the same functionality as the **OpenTextFile** method of the **FileSystemObject**. In addition, the **OpenAsTextStream** method can be used to write to a file.
The following code illustrates the use of the  **OpenAsTextStream** method:



```vb
Sub TextStreamTest
    Const ForReading = 1, ForWriting = 2, ForAppending = 3
    Const TristateUseDefault = -2, TristateTrue = -1, TristateFalse = 0
    Dim fs, f, ts, s
    Set fs = CreateObject("Scripting.FileSystemObject")
    fs.CreateTextFile "test1.txt"            'Create a file
    Set f = fs.GetFile("test1.txt")
    Set ts = f.OpenAsTextStream(ForWriting, TristateUseDefault)
    ts.Write "Hello World"
    ts.Close
    Set ts = f.OpenAsTextStream(ForReading, TristateUseDefault)
    s = ts.ReadLine
    MsgBox s
    ts.Close
End Sub
```


