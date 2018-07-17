---
title: AddressList.IsReadOnly Property (Outlook)
keywords: vbaol11.chm2030
f1_keywords:
- vbaol11.chm2030
ms.prod: outlook
api_name:
- Outlook.AddressList.IsReadOnly
ms.assetid: 45d40efc-08c0-e2d7-572a-a5e60efb7d2f
ms.date: 06/08/2017
---


# AddressList.IsReadOnly Property (Outlook)

Returns a  **Boolean** value that indicates that the **[AddressList](addresslist-object-outlook.md)** object cannot be modified. Read-only.


## Syntax

 _expression_ . **IsReadOnly**

 _expression_ A variable that represents an **AddressList** object.


## Remarks

The  **IsReadOnly** property refers to adding and deleting the entries in the address book container represented by the **AddressList** object. The property is **True** if no entries can be added or deleted. The property is **False** if the container can be modified, that is, if address entries can be added to and deleted from the container. It refers to the address book entries in the context of the address book container. It does not indicate whether the contents of the individual entries themselves can be modified.


## See also


#### Concepts


[AddressList Object](addresslist-object-outlook.md)

