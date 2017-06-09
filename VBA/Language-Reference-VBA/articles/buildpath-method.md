---
title: BuildPath Method
keywords: vblr6.chm2182031
f1_keywords:
- vblr6.chm2182031
ms.prod: office
api_name:
- Office.BuildPath
ms.assetid: 55f3dbad-0e0a-1968-a749-fe87986e9690
ms.date: 06/08/2017
---


# BuildPath Method



 **Description**
Appends a name to an existing path.
 **Syntax**
 _object_. **BuildPath(**_path_, _name_**)**
The  **BuildPath** method syntax has these parts:


|**Part**|**Description**|
|:-----|:-----|
| _object_|Required. Always the name of a  **FileSystemObject**.|
| _path_|Required. Existing path to which  _name_ is appended. Path can be absolute or relative and need not specify an existing folder.|
| _name_|Required. Name being appended to the existing  _path_.|
 **Remarks**
The  **BuildPath** method inserts an additional path separator between the existing path and the name, only if necessary.

