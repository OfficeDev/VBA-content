---
title: GetFile Method
keywords: vblr6.chm2182054
f1_keywords:
- vblr6.chm2182054
ms.prod: office
api_name:
- Office.GetFile
ms.assetid: bdb2737e-7836-4dac-9216-6f1bd8f92aa8
ms.date: 06/08/2017
---


# GetFile Method



 **Description**
Returns a  **File** object corresponding to the file in a specified path.
 **Syntax**
 _object_. **GetFile(**_filespec_**)**
The  **GetFile** method syntax has these parts:


|**Part**|**Description**|
|:-----|:-----|
| _object_|Required. Always the name of a  **FileSystemObject**.|
| _filespec_|Required. The  _filespec_ is the path (absolute or relative) to a specific file.|
 **Remarks**
An error occurs if the specified file does not exist.

