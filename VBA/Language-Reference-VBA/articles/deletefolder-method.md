---
title: DeleteFolder Method
keywords: vblr6.chm2182037
f1_keywords:
- vblr6.chm2182037
ms.prod: office
api_name:
- Office.DeleteFolder
ms.assetid: 2eec70c2-7558-1dd1-898a-95ea36de8d36
ms.date: 06/08/2017
---


# DeleteFolder Method



 **Description**
Deletes a specified folder and its contents.
 **Syntax**
 _object_. **DeleteFolder**_folderspec_ [, _force_ ]
The  **DeleteFolder** method syntax has these parts:


|**Part**|**Description**|
|:-----|:-----|
| _object_|Required. Always the name of a  **FileSystemObject**.|
| _folderspec_|Required. The name of the folder to delete. The  _folderspec_ can contain wildcard characters in the last path component.|
| _force_|Optional.  **Boolean** value that is **True** if folders with the read-only attribute set are to be deleted; **False** (default) if they are not.|
 **Remarks**
The  **DeleteFolder** method does not distinguish between folders that have contents and those that do not. The specified folder is deleted regardless of whether or not it has contents.
An error occurs if no matching folders are found. The  **DeleteFolder** method stops on the first error it encounters. No attempt is made to roll back or undo any changes that were made before an error occurred.

