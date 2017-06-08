---
title: BusinessCardView.Session Property (Outlook)
keywords: vbaol11.chm2919
f1_keywords:
- vbaol11.chm2919
ms.prod: outlook
api_name:
- Outlook.BusinessCardView.Session
ms.assetid: 18e5fb02-1d57-3c47-74ed-0409d734b4cb
ms.date: 06/08/2017
---


# BusinessCardView.Session Property (Outlook)

Returns the  **[NameSpace](namespace-object-outlook.md)** object for the current session. Read-only.


## Syntax

 _expression_ . **Session**

 _expression_ A variable that represents a **BusinessCardView** object.


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


[BusinessCardView Object](businesscardview-object-outlook.md)

