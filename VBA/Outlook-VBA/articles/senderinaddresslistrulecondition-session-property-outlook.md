---
title: SenderInAddressListRuleCondition.Session Property (Outlook)
keywords: vbaol11.chm2466
f1_keywords:
- vbaol11.chm2466
ms.prod: outlook
api_name:
- Outlook.SenderInAddressListRuleCondition.Session
ms.assetid: ec6ac5e8-9e21-a073-c179-7050e5a9b6c2
ms.date: 06/08/2017
---


# SenderInAddressListRuleCondition.Session Property (Outlook)

Returns the  **[NameSpace](namespace-object-outlook.md)** object for the current session. Read-only.


## Syntax

 _expression_ . **Session**

 _expression_ A variable that represents a **SenderInAddressListRuleCondition** object.


## Remarks

The  **Session** property and the **[GetNamespace](application-getnamespace-method-outlook.md)** method can be used interchangeably to obtain the **NameSpace** object for the current session. Both members serve the same purpose. For example, the following statements perform the same function:


```vb
Set objNamespace = Application.GetNamespace("MAPI") 
```


```vb
Set objSession = Application.Session
```


## See also


#### Concepts


[SenderInAddressListRuleCondition Object](senderinaddresslistrulecondition-object-outlook.md)

