---
title: GetAbsolutePathName Method
keywords: vblr6.chm2182045
f1_keywords:
- vblr6.chm2182045
ms.prod: office
api_name:
- Office.GetAbsolutePathName
ms.assetid: 49209a8f-6346-b32a-55d7-d72692b4defb
ms.date: 06/08/2017
---


# GetAbsolutePathName Method



 **Description**
Returns a complete and unambiguous path from a provided path specification.
 **Syntax**
 _object_. **GetAbsolutePathName(**_pathspec_**)**
The  **GetAbsolutePathName** method syntax has these parts:


| <strong>Part</strong> | <strong>Description</strong>                                               |
|:----------------------|:---------------------------------------------------------------------------|
| <em>object</em>       | Required. Always the name of a  <strong>FileSystemObject</strong>.         |
| <em>pathspec</em>     | Required. Path specification to change to a complete and unambiguous path. |

 <strong>Remarks</strong>
A path is complete and unambiguous if it provides a complete reference from the root of the specified drive. A complete path can only end with a path separator character ( 
<strong>\</strong> ) if it specifies the root folder of a mapped drive.
Assuming the current directory is c
:\mydocuments\reports, the following table illustrates the behavior of the  <strong>GetAbsolutePathName</strong> method.


|**_pathspec_**|**Returned path**|
|:-----|:-----|
|"c:"|"c:\mydocuments\reports"|
|"c:.."|"c:\mydocuments"|
|"c:\\\"|"c:\"|
|"c:*.*\may97"|"c:\mydocuments\reports\*.*\may97"|
|"region1"|"c:\mydocuments\reports\region1"|
|"c:\..\..\mydocuments"|"c:\mydocuments"|

