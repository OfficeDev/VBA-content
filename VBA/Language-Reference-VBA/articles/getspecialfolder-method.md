---
title: GetSpecialFolder Method
keywords: vblr6.chm2182057
f1_keywords:
- vblr6.chm2182057
ms.prod: office
api_name:
- Office.GetSpecialFolder
ms.assetid: f10f5721-43a2-6c0d-67a2-a1192c127c06
ms.date: 06/08/2017
---


# GetSpecialFolder Method



 **Description**
Returns the special folder specified.
 **Syntax**
 _object_**.GetSpecialFolder(**_folderspec_**)**
The  **GetSpecialFolder** method syntax has these parts:


|**Part**|**Description**|
|:-----|:-----|
| _object_|Required. Always the name of a  **FileSystemObject**.|
| _folderspec_|Required. The name of the special folder to be returned. Can be any of the constants shown in the Settings section.|
 **Settings**
The  _folderspec_ argument can have any of the following values:


|**Constant**|**Value**|**Description**|
|:-----|:-----|:-----|
|**WindowsFolder**|0|The Windows folder contains files installed by the Windows operating system.|
|**SystemFolder**|1|The System folder contains libraries, fonts, and device drivers.|
|**TemporaryFolder**|2|The Temp folder is used to store temporary files. Its path is found in the TMP environment variable.|

