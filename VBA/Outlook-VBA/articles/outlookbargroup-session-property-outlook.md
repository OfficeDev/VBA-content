---
title: OutlookBarGroup.Session Property (Outlook)
keywords: vbaol11.chm323
f1_keywords:
- vbaol11.chm323
ms.prod: outlook
api_name:
- Outlook.OutlookBarGroup.Session
ms.assetid: eb75d479-7217-51b3-6426-53ff960e9c60
ms.date: 06/08/2017
---


# OutlookBarGroup.Session Property (Outlook)

Returns the  **[NameSpace](namespace-object-outlook.md)** object for the current session. Read-only.


## Syntax

 _expression_ . **Session**

 _expression_ A variable that represents an **OutlookBarGroup** object.


## Remarks

The  **Session** property and the **[GetNamespace](application-getnamespace-method-outlook.md)** method can be used interchangeably to obtain the **NameSpace** object for the current session. Both members serve the same purpose. For example, the following statements do the same function:


```vb
Set objNamespace = Application.GetNamespace("MAPI") 
```


```vb
Set objSession = Application.Session
```


## See also


#### Concepts


[OutlookBarGroup Object](outlookbargroup-object-outlook.md)

