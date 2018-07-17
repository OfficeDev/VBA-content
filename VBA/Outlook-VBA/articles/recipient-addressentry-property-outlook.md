---
title: Recipient.AddressEntry Property (Outlook)
keywords: vbaol11.chm2345
f1_keywords:
- vbaol11.chm2345
ms.prod: outlook
api_name:
- Outlook.Recipient.AddressEntry
ms.assetid: 3b2b524e-4dd5-9ff4-98cc-811746ea0453
ms.date: 06/08/2017
---


# Recipient.AddressEntry Property (Outlook)

Returns the  **[AddressEntry](addressentry-object-outlook.md)** object corresponding to the resolved recipient. Read/write.


## Syntax

 _expression_ . **AddressEntry**

 _expression_ A variable that represents a **Recipient** object.


## Remarks

Accessing the  **AddressEntry** property forces resolution of an unresolved recipient name. If the name cannot be resolved, an error is returned. If the recipient is resolved, the **[Resolved](recipient-resolved-property-outlook.md)** property is **True** .


## See also


#### Concepts


[Recipient Object](recipient-object-outlook.md)

