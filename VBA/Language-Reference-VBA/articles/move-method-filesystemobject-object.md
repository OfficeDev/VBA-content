---
title: Move Method (FileSystemObject object)
keywords: vblr6.chm2182006
f1_keywords:
- vblr6.chm2182006
ms.prod: office
ms.assetid: 9191e310-2b92-fd13-f04a-e34ca2743b7e
ms.date: 06/08/2017
---


# Move Method (FileSystemObject object)



 **Description**
Moves a specified file or folder from one location to another.
 **Syntax**
 _object_. **Move**_destination_
The  **Move** method syntax has these parts:


|**Part**|**Description**|
|:-----|:-----|
| _object_|Required. Always the name of a  **File** or **Folder** object.|
| _destination_|Required. Destination where the file or folder is to be moved. Wildcard characters are not allowed.|
 **Remarks**
The results of the  **Move** method on a **File** or **Folder** are identical to operations performed using **FileSystemObject.MoveFile** or **FileSystemObject.MoveFolder**. You should note, however, that the alternative methods are capable of moving multiple files or folders.

