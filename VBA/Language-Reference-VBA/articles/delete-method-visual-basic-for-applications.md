---
title: Delete Method (Visual Basic for Applications)
keywords: vblr6.chm2182005
f1_keywords:
- vblr6.chm2182005
ms.prod: office
ms.assetid: 698cb2bd-17b2-2560-f406-09bb9991b86c
ms.date: 06/08/2017
---


# Delete Method (Visual Basic for Applications)



 **Description**
Deletes a specified file or folder.
 **Syntax**
 _object_. **Delete**_force_
The  **Delete** method syntax has these parts:


|**Part**|**Description**|
|:-----|:-----|
| _object_|Required. Always the name of a  **File** or **Folder** object.|
| _force_|Optional.  **Boolean** value that is **True** if files or folders with the read-only attribute set are to be deleted; **False** (default) if they are not.|
 **Remarks**
An error occurs if the specified file or folder does not exist.
The results of the  **Delete** method on a **File** or **Folder** are identical to operations performed using **FileSystemObject.DeleteFile** or **FileSystemObject.DeleteFolder**.
The  **Delete** method does not distinguish between folders that have contents and those that do not. The specified folder is deleted regardless of whether or not it has contents.

