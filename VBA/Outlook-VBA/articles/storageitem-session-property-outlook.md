---
title: StorageItem.Session Property (Outlook)
keywords: vbaol11.chm2139
f1_keywords:
- vbaol11.chm2139
ms.prod: outlook
api_name:
- Outlook.StorageItem.Session
ms.assetid: e3a005d0-daa3-853b-e603-c084ffb5d1db
ms.date: 06/08/2017
---


# StorageItem.Session Property (Outlook)

Returns the  **[NameSpace](namespace-object-outlook.md)** object for the current session. Read-only.


## Syntax

 _expression_ . **Session**

 _expression_ A variable that represents a **StorageItem** object.


## Remarks

Returns  **Null** ( **Nothing** in Visual Basic) if there is no logged-on session.

The  **Session** property and the **[GetNamespace](application-getnamespace-method-outlook.md)** method can be used interchangeably to obtain the **NameSpace** object for the current session. Both members serve the same purpose. For example, the following statements perform the same function:




```vb
Set objNamespace = Application.GetNamespace("MAPI") 
```




```vb
Set objSession = Application.Session
```


## See also


#### Concepts


[StorageItem Object](storageitem-object-outlook.md)

