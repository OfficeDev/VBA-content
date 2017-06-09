---
title: Designer Property (VBA Add-In Object Model)
keywords: vbob6.chm102191
f1_keywords:
- vbob6.chm102191
ms.prod: office
ms.assetid: 4ec4b33f-35c5-c5b6-554a-464c068588ff
ms.date: 06/08/2017
---


# Designer Property (VBA Add-In Object Model)



Returns the object that enables you to access the design characteristics of a component.
 **Remarks**
If the object has an open [designer](vbe-glossary.md), the  **Designer** property returns the open designer; otherwise a new designer is created. The designer is a characteristic of certain **VBComponent** objects. For example, when you create certain types of **VBComponent** object, a designer is created along with the object. A component can have only one designer, and it's always the same designer. The **Designer** property enables you to access a component-specific object. In some cases, such as in[standard modules](vbe-glossary.md) and[class modules](vbe-glossary.md), a designer isn't created because that type of  **VBComponent** object doesn't support a designer.
The  **Designer** property returns **Nothing** if the **VBComponent** object doesn't have a designer.

