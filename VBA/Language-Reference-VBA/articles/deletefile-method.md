---
title: DeleteFile Method
keywords: vblr6.chm2182036
f1_keywords:
- vblr6.chm2182036
ms.prod: office
api_name:
- Office.DeleteFile
ms.assetid: e036b009-4fd9-297a-de24-acc0dbc96c7a
ms.date: 06/08/2017
---


# DeleteFile Method



 **Description**
Deletes a specified file.
 **Syntax**
 _object_. **DeleteFile**_filespec_ [, _force_ ]
The  **DeleteFile** method syntax has these parts:


| <strong>Part</strong> | <strong>Description</strong>                                                                                                                                                           |
|:----------------------|:---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------|
| <em>object</em>       | Required. Always the name of a  <strong>FileSystemObject</strong>.                                                                                                                     |
| <em>filespec</em>     | Required. The name of the file to delete. The  <em>filespec</em> can contain wildcard characters in the last path component.                                                           |
| <em>force</em>        | Optional.  <strong>Boolean</strong> value that is <strong>True</strong> if files with the read-only attribute set are to be deleted; <strong>False</strong> (default) if they are not. |

 **Remarks**
An error occurs if no matching files are found. The  **DeleteFile** method stops on the first error it encounters. No attempt is made to roll back or undo any changes that were made before an error occurred.

