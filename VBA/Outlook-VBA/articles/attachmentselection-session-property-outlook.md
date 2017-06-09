---
title: AttachmentSelection.Session Property (Outlook)
keywords: vbaol11.chm2941
f1_keywords:
- vbaol11.chm2941
ms.prod: outlook
api_name:
- Outlook.AttachmentSelection.Session
ms.assetid: cd03fab6-60bd-5e71-3053-b6cc61fda58f
ms.date: 06/08/2017
---


# AttachmentSelection.Session Property (Outlook)

Returns the  **[NameSpace](namespace-object-outlook.md)** object for the current session. Read-only.


## Syntax

 _expression_ . **Session**

 _expression_ A variable that represents an **AttachmentSelection** object.


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


[AttachmentSelection Object](attachmentselection-object-outlook.md)

