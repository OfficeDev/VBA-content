---
title: Pages.Session Property (Outlook)
keywords: vbaol11.chm393
f1_keywords:
- vbaol11.chm393
ms.prod: outlook
api_name:
- Outlook.Pages.Session
ms.assetid: cebf5807-8f1f-05f4-e990-35fb00e07f0a
ms.date: 06/08/2017
---


# Pages.Session Property (Outlook)

Returns the  **[NameSpace](namespace-object-outlook.md)** object for the current session. Read-only.


## Syntax

 _expression_ . **Session**

 _expression_ A variable that represents a **Pages** object.


## Remarks

The  **Session** property and the **[GetNamespace](application-getnamespace-method-outlook.md)** method can be used interchangeably to obtain the **NameSpace** object for the current session. Both members serve the same purpose. For example, the following statements do the same function:


```vb
Set objNamespace = Application.GetNamespace("MAPI") 
```


```vb
Set objSession = Application.Session
```


## See also


#### Other resources


[Pages Object](http://msdn.microsoft.com/library/20a5339d-1dc7-9b61-d725-d13db72c5f65%28Office.15%29.aspx)


