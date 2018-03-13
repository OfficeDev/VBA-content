---
title: Copy Method (Visual Basic for Applications)
keywords: vblr6.chm2182004
f1_keywords:
- vblr6.chm2182004
ms.prod: office
ms.assetid: 3477c158-643a-5e29-e4c2-b451e8603542
ms.date: 06/08/2017
---


# Copy Method (Visual Basic for Applications)



 **Description**
Copies a specified file or folder from one location to another.
 **Syntax**
 _object_. **Copy**_destination_ [, _overwrite_ ]
The  **Copy** method syntax has these parts:


| <strong>Part</strong> | <strong>Description</strong>                                                                                                                                                  |
|:----------------------|:------------------------------------------------------------------------------------------------------------------------------------------------------------------------------|
| <em>object</em>       | Required. Always the name of a  <strong>File</strong> or <strong>Folder</strong> object.                                                                                      |
| <em>destination</em>  | Required. Destination where the file or folder is to be copied. Wildcard characters are not allowed.                                                                          |
| <em>overwrite</em>    | Optional.  <strong>Boolean</strong> value that is <strong>True</strong> (default) if existing files or folders are to be overwritten; <strong>False</strong> if they are not. |

 **Remarks**
The results of the  **Copy** method on a **File** or **Folder** are identical to operations performed using **FileSystemObject.CopyFile** or **FileSystemObject.CopyFolder** where the file or folder referred to by _object_ is passed as an argument. You should note, however, that the alternative methods are capable of copying multiple files or folders.

