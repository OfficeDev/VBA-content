---
title: FolderExists Method
keywords: vblr6.chm2182042
f1_keywords:
- vblr6.chm2182042
ms.prod: office
api_name:
- Office.FolderExists
ms.assetid: 5a4e9c53-7561-3065-f2b3-545e9efc503d
ms.date: 06/08/2017
---


# FolderExists Method



 **Description**
Returns  **True** if a specified folder exists; **False** if it does not.
 **Syntax**
 _object_. **FolderExists(**_folderspec_ )
The  **FolderExists** method syntax has these parts:


|**Part**|**Description**|
|:-----|:-----|
| _object_|Required. Always the name of a  **FileSystemObject**.|
| _folderspec_|Required. The name of the folder whose existence is to be determined. A complete path specification (either absolute or relative) must be provided if the folder isn't expected to exist in the current folder.|

