---
title: Folder.Store Property (Outlook)
keywords: vbaol11.chm2016
f1_keywords:
- vbaol11.chm2016
ms.prod: outlook
api_name:
- Outlook.Folder.Store
ms.assetid: 347d3031-01cf-a248-4abc-f749feb811a4
ms.date: 06/08/2017
---


# Folder.Store Property (Outlook)

Returns a  **[Store](store-object-outlook.md)** object representing the store that contains the **[Folder](folder-object-outlook.md)** object. Read-only.


## Syntax

 _expression_ . **Store**

 _expression_ A variable that represents a **Folder** object.


## Remarks

This property returns a  **Store** object except in the case where the **Folder** is a shared folder (returned by **[NameSpace.GetSharedDefaultFolder](namespace-getshareddefaultfolder-method-outlook.md)** ). In this case, one user has delegated access to a default folder to another user; a call to **Folder.Store** will return **Null** .


## See also


#### Concepts


[Folder Object](folder-object-outlook.md)

