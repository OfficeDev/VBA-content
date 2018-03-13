---
title: OpenTextFile Method
keywords: vblr6.chm2182061
f1_keywords:
- vblr6.chm2182061
ms.prod: office
api_name:
- Office.OpenTextFile
ms.assetid: f44f7bc5-e48b-05f2-eb22-5b02701d449e
ms.date: 06/08/2017
---


# OpenTextFile Method



 **Description**
Opens a specified file and returns a  **TextStream** object that can be used to read from or append to the file.
 **Syntax**
 _object_. **OpenTextFile(**_filename_ [ **,**_iomode_ [ **,**_create_ [ **,**_format_ ]]] **)**
The  **OpenTextFile** method has these parts:


| <strong>Part</strong> | <strong>Description</strong>                                                                                                                                                                                                                                                                 |
|:----------------------|:---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------|
| <em>object</em>       | Required. Always the name of a  <strong>FileSystemObject</strong>.                                                                                                                                                                                                                           |
| <em>filename</em>     | Required. [String expression](vbe-glossary.md) that identifies the file to open.                                                                                                                                                                                                             |
| <em>iomode</em>       | Optional. Indicates input/output mode. Can be one of two constants, either  <strong>ForReading</strong> or <strong>ForAppending</strong>.                                                                                                                                                    |
| <em>create</em>       | Optional.  <strong>Boolean</strong> value that indicates whether a new file can be created if the specified <em>filename</em> doesn't exist. The value is <strong>True</strong> if a new file is created; <strong>False</strong> if it isn't created. The default is <strong>False</strong>. |
| <em>format</em>       | Optional. One of three  <strong>Tristate</strong> values used to indicate the format of the opened file. If omitted, the file is opened as ASCII.                                                                                                                                            |

 **Settings**
The  _iomode_ argument can have either of the following settings:


| <strong>Constant</strong>     | <strong>Value</strong> | <strong>Description</strong>                                |
|:------------------------------|:-----------------------|:------------------------------------------------------------|
| <strong>ForReading</strong>   | 1                      | Open a file for reading only. You can't write to this file. |
| <strong>ForAppending</strong> | 8                      | Open a file and write to the end of the file.               |

The  _format_ argument can have any of the following settings:


| <strong>Constant</strong>           | <strong>Value</strong> | <strong>Description</strong>             |
|:------------------------------------|:-----------------------|:-----------------------------------------|
| <strong>TristateUseDefault</strong> | -2                     | Opens the file using the system default. |
| <strong>TristateTrue</strong>       | -1                     | Opens the file as Unicode.               |
| <strong>TristateFalse</strong>      | 0                      | Opens the file as ASCII.                 |

 **Remarks**
The following code illustrates the use of the  **OpenTextFile** method to open a file for appending text:



```vb
Sub OpenTextFileTest
    Const ForReading = 1, ForWriting = 2, ForAppending = 3
    Dim fs, f
    Set fs = CreateObject("Scripting.FileSystemObject")
    Set f = fs.OpenTextFile("c:\testfile.txt", ForAppending,TristateFalse)
    f.Write "Hello world!"
    f.Close
End Sub
```


