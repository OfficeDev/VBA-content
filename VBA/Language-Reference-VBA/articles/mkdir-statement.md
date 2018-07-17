---
title: MkDir Statement
keywords: vblr6.chm1008975
f1_keywords:
- vblr6.chm1008975
ms.prod: office
ms.assetid: b79fdad3-a1c2-7af3-c679-09d35d4b0d87
ms.date: 06/08/2017
---


# MkDir Statement

Creates a new directory or folder.

 **Syntax**

 **MkDir** _path_

The required  _path_ [argument](vbe-glossary.md) is a [string expression](vbe-glossary.md) that identifies the directory or folder to be created. The _path_ may include the drive. If no drive is specified, **MkDir** creates the new directory or folder on the current drive. If the directory or folder already exists, **MkDir** will fail with run-time error 75 (Path/File access error.)

## Example

This example uses the  **MkDir** statement to create a directory or folder. If the drive is not specified, the new directory or folder is created on the current drive.


```
MkDir "MYDIR" ' Make new directory or folder. 

```


