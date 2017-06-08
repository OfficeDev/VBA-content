---
title: TableOfAuthoritiesCategory.Creator Property (Word)
keywords: vbawd10.chm158794729
f1_keywords:
- vbawd10.chm158794729
ms.prod: word
api_name:
- Word.TableOfAuthoritiesCategory.Creator
ms.assetid: cf7f8dec-174a-f10d-f631-c65ad589f97f
ms.date: 06/08/2017
---


# TableOfAuthoritiesCategory.Creator Property (Word)

Returns a 32-bit integer that indicates the application in which the specified object was created. Read-only  **Long** .


## Syntax

 _expression_ . **Creator**

 _expression_ Required. A variable that represents a **[TableOfAuthoritiesCategory](tableofauthoritiescategory-object-word.md)** object.


## Remarks

If the object was created in Microsoft Word, the  **Creator** property returns the hexadecimal number 4D535744, which represents the string "MSWD." This property was primarily designed to be used on the Macintosh, where each application has a four-character creator code. For example, Microsoft Word has the creator code MSWD. For additional information about this property, consult the language reference Help included with Microsoft Office Macintosh Edition.


## See also


#### Concepts


[TableOfAuthoritiesCategory Object](tableofauthoritiescategory-object-word.md)

