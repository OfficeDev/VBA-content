---
title: ReportItem.Session Property (Outlook)
keywords: vbaol11.chm1636
f1_keywords:
- vbaol11.chm1636
ms.prod: outlook
api_name:
- Outlook.ReportItem.Session
ms.assetid: b9599afe-1c2b-36b2-2ce4-8e781f32975a
ms.date: 06/08/2017
---


# ReportItem.Session Property (Outlook)

Returns the  **[NameSpace](namespace-object-outlook.md)** object for the current session. Read-only.


## Syntax

 _expression_ . **Session**

 _expression_ A variable that represents a **ReportItem** object.


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


[ReportItem Object](reportitem-object-outlook.md)

