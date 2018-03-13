---
title: GetParentFolderName Method
keywords: vblr6.chm2182056
f1_keywords:
- vblr6.chm2182056
ms.prod: office
api_name:
- Office.GetParentFolderName
ms.assetid: 445e969a-6a01-6cb0-aff7-378717277c69
ms.date: 06/08/2017
---


# GetParentFolderName Method



 **Description**
Returns a string containing the name of the parent folder of the last component in a specified path.
 **Syntax**
 _object_. **GetParentFolderName(**_path_**)**
The  **GetParentFolderName** method syntax has these parts:


| <strong>Part</strong> | <strong>Description</strong>                                                                   |
|:----------------------|:-----------------------------------------------------------------------------------------------|
| <em>object</em>       | Required. Always the name of a  <strong>FileSystemObject</strong>.                             |
| <em>path</em>         | Required. The path specification for the component whose parent folder name is to be returned. |

 **Remarks**
The  **GetParentFolderName** method returns a zero-length string ("") if there is no parent folder for the component specified in the _path_ argument.

 **Note**  The  **GetParentFolderName** method works only on the provided _path_ string. It does not attempt to resolve the path, nor does it check for the existence of the specified path.


