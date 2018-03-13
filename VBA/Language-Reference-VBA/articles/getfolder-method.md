---
title: GetFolder Method
keywords: vblr6.chm2182055
f1_keywords:
- vblr6.chm2182055
ms.prod: office
api_name:
- Office.GetFolder
ms.assetid: 772f1ae7-ac29-d4b4-e08a-d8553375510d
ms.date: 06/08/2017
---


# GetFolder Method



 **Description**
Returns a  **Folder** object corresponding to the folder in a specified path.
 **Syntax**
 _object_. **GetFolder(**_folderspec_**)**
The  **GetFolder** method syntax has these parts:


| <strong>Part</strong> | <strong>Description</strong>                                                                |
|:----------------------|:--------------------------------------------------------------------------------------------|
| <em>object</em>       | Required. Always the name of a  <strong>FileSystemObject</strong>.                          |
| <em>folderspec</em>   | Required. The  <em>folderspec</em> is the path (absolute or relative) to a specific folder. |

 **Remarks**
An error occurs if the specified folder does not exist.

