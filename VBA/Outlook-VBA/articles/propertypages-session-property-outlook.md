---
title: PropertyPages.Session Property (Outlook)
keywords: vbaol11.chm163
f1_keywords:
- vbaol11.chm163
ms.prod: outlook
api_name:
- Outlook.PropertyPages.Session
ms.assetid: 0a6c6235-b27b-72d4-bd17-c94627b91d41
ms.date: 06/08/2017
---


# PropertyPages.Session Property (Outlook)

Returns the  **[NameSpace](namespace-object-outlook.md)** object for the current session. Read-only.


## Syntax

 _expression_ . **Session**

 _expression_ A variable that represents a **PropertyPages** object.


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


[PropertyPages Object](propertypages-object-outlook.md)

