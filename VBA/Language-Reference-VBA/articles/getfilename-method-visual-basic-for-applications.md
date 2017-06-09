---
title: GetFileName Method (Visual Basic for Applications)
keywords: vblr6.chm2182053
f1_keywords:
- vblr6.chm2182053
ms.prod: office
ms.assetid: af5ca68f-ec3e-409c-dcb4-75202169ccb8
ms.date: 06/08/2017
---


# GetFileName Method (Visual Basic for Applications)



 **Description**
Returns the last component of specified path that is not part of the drive specification.
 **Syntax**
 _object_. **GetFileName(**_pathspec_**)**
The  **GetFileName** method syntax has these parts:


|**Part**|**Description**|
|:-----|:-----|
| _object_|Required. Always the name of a  **FileSystemObject**.|
| _pathspec_|Required. The path (absolute or relative) to a specific file.|
 **Remarks**
The  **GetFileName** method returns a zero-length string ("") if _pathspec_ does not end with the named component.

 **Note**  The  **GetFileName** method works only on the provided path string. It does not attempt to resolve the path, nor does it check for the existence of the specified path.


