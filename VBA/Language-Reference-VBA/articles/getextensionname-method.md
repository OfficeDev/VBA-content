---
title: GetExtensionName Method
keywords: vblr6.chm2182052
f1_keywords:
- vblr6.chm2182052
ms.prod: office
api_name:
- Office.GetExtensionName
ms.assetid: 0fa9da71-7938-c50c-6fed-8a23d6a680d1
ms.date: 06/08/2017
---


# GetExtensionName Method



 **Description**
Returns a string containing the extension name for the last component in a path.
 **Syntax**
 _object_. **GetExtensionName(**_path_**)**
The  **GetExtensionName** method syntax has these parts:


|**Part**|**Description**|
|:-----|:-----|
| _object_|Required. Always the name of a  **FileSystemObject**.|
| _path_|Required. The path specification for the component whose extension name is to be returned.|
 **Remarks**
For network drives, the root directory ( **\** ) is considered to be a component.
The  **GetExtensionName** method returns a zero-length string ("") if no component matches the _path_ argument.

