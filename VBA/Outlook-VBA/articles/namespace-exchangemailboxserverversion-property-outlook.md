---
title: NameSpace.ExchangeMailboxServerVersion Property (Outlook)
keywords: vbaol11.chm792
f1_keywords:
- vbaol11.chm792
ms.prod: outlook
api_name:
- Outlook.NameSpace.ExchangeMailboxServerVersion
ms.assetid: 01e83a30-f574-1ff6-34de-85c14ecc09c1
ms.date: 06/08/2017
---


# NameSpace.ExchangeMailboxServerVersion Property (Outlook)

Returns a  **String** value that represents the full version number of the Exchange server that hosts the primary Exchange account mailbox. Read-only.


## Syntax

 _expression_ . **ExchangeMailboxServerVersion**

 _expression_ An expression that returns a **[NameSpace](namespace-object-outlook.md)** object.


## Remarks

This property returns a string that contains the version number of the Exchange server for the active mailbox. The version number has the following four parts.


```
<major version>.<minor version>.<build number>.<revision>
```

Not all parts may be present in the version number, depending on the version information that is supplied by the Microsoft Exchange Server. For example, this property returns "6.5.7638" for Microsoft Exchange Server 2003 Service Pack 2.

If an Exchange mailbox is not present in the namespace, this property returns an empty string.


## See also


#### Concepts


[NameSpace Object](namespace-object-outlook.md)

