---
title: DeleteLines Method (VBA Add-In Object Model)
keywords: vbob6.chm104014
f1_keywords:
- vbob6.chm104014
ms.prod: office
ms.assetid: b6e1bd5d-23b2-0bc4-bcc6-b7e371df4b93
ms.date: 06/08/2017
---


# DeleteLines Method (VBA Add-In Object Model)



Deletes a single line or a specified range of lines.
 **Syntax**
 _object_**.DeleteLines (**_startline_ [, _count_ ] **)**
The  **DeleteLines** syntax has these parts:


|**Part**|**Description**|
|:-----|:-----|
| _object_|Required. An [object expression](vbe-glossary.md) that evaluates to an object in the Applies To list.|
| _startline_|Required. A [Long](vbe-glossary.md) specifying the first line you want to delete.|
| _count_|Optional. A  **Long** specifying the number of lines you want to delete.|
 **Remarks**
If you don't specify how many lines you want to delete,  **DeleteLines** deletes one line.

