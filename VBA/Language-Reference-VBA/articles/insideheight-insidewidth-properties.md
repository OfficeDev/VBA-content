---
title: InsideHeight, InsideWidth Properties
keywords: fm20.chm5225045
f1_keywords:
- fm20.chm5225045
ms.prod: office
ms.assetid: 8db4373d-0807-ec2a-f9df-37ebcbf8ef47
ms.date: 06/08/2017
---


# InsideHeight, InsideWidth Properties



 **InsideHeight** returns the height, in[points](vbe-glossary.md), of the [client region](glossary-vba.md) inside a form. **InsideWidth** returns the width, in points, of the client region inside a form.
 **Syntax**
 _object_. **InsideHeight**
 _object_. **InsideWidth**
The  **InsideHeight** and **InsideWidth** property syntaxes have these parts:


|**Part**|**Description**|
|:-----|:-----|
| _object_|Required. A valid object.|
 **Remarks**
The  **InsideHeight** and **InsideWidth** properties are read-only. If the region includes a scroll bar, the returned value does not include the height or width of the scroll bar.

