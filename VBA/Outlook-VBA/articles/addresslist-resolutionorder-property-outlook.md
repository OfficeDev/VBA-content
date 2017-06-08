---
title: AddressList.ResolutionOrder Property (Outlook)
keywords: vbaol11.chm2034
f1_keywords:
- vbaol11.chm2034
ms.prod: outlook
api_name:
- Outlook.AddressList.ResolutionOrder
ms.assetid: e92bd83f-349b-d6e7-a5fb-7a6d893406a0
ms.date: 06/08/2017
---


# AddressList.ResolutionOrder Property (Outlook)

Returns an  **Integer** that represents the order of this **[AddressList](addresslist-object-outlook.md)** in the custom scroll list in the **Addressing** dialog box. Read-only.


## Syntax

 _expression_ . **ResolutionOrder**

 _expression_ A variable that represents an **AddressList** object.


## Remarks

The  **ResolutionOrder** property corresponds to the position of the **AddressList** in the list box which is labeled **When sending mail, check address lists in this order** in the **Addressing** dialog box and is available by clicking **Tools**, and then  **Options** in the **Address Book** dialog box. Note that this behavior is independent of whether **Custom** is selected in the **Addressing** dialog box. For example, if **Start with Global Address List** is selected and the scroll list displays **Contacts** as the first item, the **ResolutionOrder** property of the **Contacts** address list is still 1, even though when resolving recipients, Outlook uses the Global Address List first.

When  **Custom** is selected in the **Addressing** dialog box, the **ResolutionOrder** property reflects the order that Outlook uses when resolving recipient names. The value of **ResolutionOrder** is one-based. In this case, the first address list to be used for resolving recipient names has a **ResolutionOrder** equal to 1, the second address list 2, and so forth. If an address list is not used to resolve addresses, then its **ResolutionOrder** has a value of -1. Programmatically, when **Custom** is selected in the **Addressing** dialog box, the **ResolutionOrder** property reflects the actual resolution order that **[Recipients.ResolveAll](recipients-resolveall-method-outlook.md)** or **[Recipient.Resolve](recipient-resolve-method-outlook.md)** uses.


## See also


#### Concepts


[AddressList Object](addresslist-object-outlook.md)

