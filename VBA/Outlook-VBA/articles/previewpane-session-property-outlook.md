---
title: PreviewPane.Session Property (Outlook)
keywords: vbaol11.chm3636
f1_keywords:
- vbaol11.chm3636
ms.assetid: 54509e05-d255-b96e-f037-14282791ea55
ms.date: 06/08/2017
ms.prod: outlook
---


# PreviewPane.Session Property (Outlook)

Returns the [NameSpace](namespace-object-outlook.md) for the current session. Read-only.


## Syntax

 _expression_ . **Session**

 _expression_ A variable that represents a **PreviewPane** object.


## Remarks

The  **Session** property and the[GetNamespace](application-getnamespace-method-outlook.md) method can be used interchangeably to obtain the **NameSpace** object for the current session. Both members serve the same purpose. For example, the following statements do the same function:

 `Set objNamespace = Application.Getnamespace("MAPI")`

 `SetjobSession = Application.Session`


## See also


#### Other resources


[PreviewPane Object (Outlook)](previewpane-object-outlook.md)


