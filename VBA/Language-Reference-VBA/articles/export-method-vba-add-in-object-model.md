---
title: Export Method (VBA Add-In Object Model)
keywords: vbob6.chm102194
f1_keywords:
- vbob6.chm102194
ms.prod: office
ms.assetid: 46cab37a-4390-219c-68f8-05cbb59c0450
ms.date: 06/08/2017
---


# Export Method (VBA Add-In Object Model)



Saves a component as a separate file or files.
 **Syntax**
 _object_**.Export(**_filename_**)**
The  **Export** syntax has these parts:


|**Part**|**Description**|
|:-----|:-----|
| _object_|Required. An [object expression](vbe-glossary.md) that evaluates to an object in the Applies To list.|
| _filename_|Required. A [String](vbe-glossary.md) specifying the name of the file that you want to export the component to.|
 **Remarks**
When you use the  **Export** method to save a component as a separate file or files, use a file name that doesn't already exist; otherwise, an error occurs.

