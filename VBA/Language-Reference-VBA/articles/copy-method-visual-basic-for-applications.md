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


|**Part**|**Description**|
|:-----|:-----|
| _object_|Required. Always the name of a  **File** or **Folder** object.|
| _destination_|Required. Destination where the file or folder is to be copied. Wildcard characters are not allowed.|
| _overwrite_|Optional.  **Boolean** value that is **True** (default) if existing files or folders are to be overwritten; **False** if they are not.|
 **Remarks**
The results of the  **Copy** method on a **File** or **Folder** are identical to operations performed using **FileSystemObject.CopyFile** or **FileSystemObject.CopyFolder** where the file or folder referred to by _object_ is passed as an argument. You should note, however, that the alternative methods are capable of copying multiple files or folders.

