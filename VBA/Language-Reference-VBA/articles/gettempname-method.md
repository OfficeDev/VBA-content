---
title: GetTempName Method
keywords: vblr6.chm2182058
f1_keywords:
- vblr6.chm2182058
ms.prod: MULTIPLEPRODUCTS
api_name:
- Office.GetTempName
ms.assetid: 43d8a9b2-b8ea-3ef8-f0ea-84ddb5467f0a
---


# GetTempName Method



 **Description**
Returns a randomly generated temporary file or folder name that is useful for performing operations that require a temporary file or folder.
 **Syntax**
 _object_. **GetTempName**
The optional  _object_ is always the name of a **FileSystemObject**.
 **Remarks**
The  **GetTempName** method does not create a file. It provides only a temporary file name that can be used with **CreateTextFile** to create a file.

