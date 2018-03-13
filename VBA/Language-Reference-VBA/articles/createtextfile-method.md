---
title: CreateTextFile Method
keywords: vblr6.chm2182035
f1_keywords:
- vblr6.chm2182035
ms.prod: office
api_name:
- Office.CreateTextFile
ms.assetid: be538862-92a8-0386-ea4f-1809fc465cb9
ms.date: 06/08/2017
---


# CreateTextFile Method



 **Description**
Creates a specified file name and returns a  **TextStream** object that can be used to read from or write to the file.
 **Syntax**
 _object_. **CreateTextFile(**_filename_ [ **,**_overwrite_ [ **,**_unicode_ ]] **)**
The  **CreateTextFile** method has these parts:


| <strong>Part</strong> | <strong>Description</strong>                                                                                                                                                                                                                                                                 |
|:----------------------|:---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------|
| <em>object</em>       | Required. Always the name of a  <strong>FileSystemObject</strong> or <strong>Folder</strong> object.                                                                                                                                                                                         |
| <em>filename</em>     | Required. [String expression](vbe-glossary.md) that identifies the file to create.                                                                                                                                                                                                           |
| <em>overwrite</em>    | Optional.  <strong>Boolean</strong> value that indicates if an existing file can be overwritten. The value is <strong>True</strong> if the file can be overwritten; <strong>False</strong> if it can't be overwritten. If omitted, existing files are not overwritten.                       |
| <em>unicode</em>      | Optional.  <strong>Boolean</strong> value that indicates whether the file is created as a Unicode or ASCII file. The value is <strong>True</strong> if the file is created as a Unicode file; <strong>False</strong> if it's created as an ASCII file. If omitted, an ASCII file is assumed. |

 **Remarks**
The following code illustrates how to use the  **CreateTextFile** method to create and open a text file:
If the  _overwrite_ argument is **False**, or is not provided, for a _filename_ that already exists, an error occurs.



```vb
Sub CreateAfile
    Set fs = CreateObject("Scripting.FileSystemObject")
    Set a = fs.CreateTextFile("c:\testfile.txt", True)
    a.WriteLine("This is a test.")
    a.Close
End Sub
```


