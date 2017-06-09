---
title: RmDir Statement
keywords: vblr6.chm1009007
f1_keywords:
- vblr6.chm1009007
ms.prod: office
ms.assetid: 7bc350d2-7d1a-7c8c-95a8-8dbf5c8f7953
ms.date: 06/08/2017
---


# RmDir Statement

Removes an existing directory or folder.

 **Syntax**

 **RmDir**_path_

The required  _path_[argument](vbe-glossary.md) is a[string expression](vbe-glossary.md) that identifies the directory or folder to be removed. The _path_ may include the drive. If no drive is specified, **RmDir** removes the directory or folder on the current drive.
 **Remarks**
An error occurs if you try to use  **RmDir** on a directory or folder containing files. Use the **Kill** statement to delete all files before attempting to remove a directory or folder.

## Example

This example uses the  **RmDir** statement to remove an existing directory or folder.


```vb
' Assume that MYDIR is an empty directory or folder. 
RmDir "MYDIR" ' Remove MYDIR. 

```


