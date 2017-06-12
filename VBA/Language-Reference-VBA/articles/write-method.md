---
title: Write Method
keywords: vblr6.chm2182081
f1_keywords:
- vblr6.chm2182081
ms.prod: office
api_name:
- Office.Write
ms.assetid: fd66062a-aa05-15a3-d88c-34a0c033f496
ms.date: 06/08/2017
---


# Write Method



 **Description**
Writes a specified string to a  **TextStream** file.
 **Syntax**
 _object_. **Write(**_string_**)**
The  **Write** method syntax has these parts:


|**Part**|**Description**|
|:-----|:-----|
| _object_|Required. Always the name of a  **TextStream** object.|
| _string_|Required. The text you want to write to the file.|
 **Remarks**
Specified strings are written to the file with no intervening spaces or characters between each string. Use the  **WriteLine** method to write a newline character or a string that ends with a newline character.

