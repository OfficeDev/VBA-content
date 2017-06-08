---
title: GetBaseName Method
keywords: vblr6.chm2182047
f1_keywords:
- vblr6.chm2182047
ms.prod: office
api_name:
- Office.GetBaseName
ms.assetid: 2f3af3ff-a996-e2f7-0048-1f5aa891d674
ms.date: 06/08/2017
---


# GetBaseName Method



 **Description**
Returns a string containing the base name of the last component, less any file extension, in a path.
 **Syntax**
 _object_. **GetBaseName(**_path_**)**
The  **GetBaseName** method syntax has these parts:


|**Part**|**Description**|
|:-----|:-----|
| _object_|Required. Always the name of a  **FileSystemObject**.|
| _path_|Required. The path specification for the component whose base name is to be returned.|
 **Remarks**
The  **GetBaseName** method returns a zero-length string ("") if no component matches the _path_ argument.

 **Note**  The  **GetBaseName** method works only on the provided _path_ string. It does not attempt to resolve the path, nor does it check for the existence of the specified path.


