---
title: TimeZones.Session Property (Outlook)
keywords: vbaol11.chm3293
f1_keywords:
- vbaol11.chm3293
ms.prod: outlook
api_name:
- Outlook.TimeZones.Session
ms.assetid: e4d6ca4d-914d-405c-8765-6ca1f97a9472
ms.date: 06/08/2017
---


# TimeZones.Session Property (Outlook)

Returns the  **[NameSpace](namespace-object-outlook.md)** object for the current session. Read-only.


## Syntax

 _expression_ . **Session**

 _expression_ A variable that represents a **TimeZones** object.


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


[TimeZones Object](timezones-object-outlook.md)

