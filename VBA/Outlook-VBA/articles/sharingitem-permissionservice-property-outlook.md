---
title: SharingItem.PermissionService Property (Outlook)
keywords: vbaol11.chm690
f1_keywords:
- vbaol11.chm690
ms.prod: outlook
api_name:
- Outlook.SharingItem.PermissionService
ms.assetid: ef50051d-420f-21db-af30-02a7d01896b6
ms.date: 06/08/2017
---


# SharingItem.PermissionService Property (Outlook)

Sets or returns an  **[OlPermissionService](olpermissionservice-enumeration-outlook.md)** constant that determines the permission service that will be used when sending a **[SharingItem](sharingitem-object-outlook.md)** protected by Information Rights Management (IRM). Read/write.


## Syntax

 _expression_ . **PermissionService**

 _expression_ A variable that represents a **SharingItem** object.


## Remarks

This property is useful only if you have more than one permission identity for a particular SMTP address. 

While you can view content that is protected by IRM on any computer running the 2007 Microsoft Office system or a later version, you must have Microsoft Office Professional Edition 2003, Microsoft Office Outlook 2007, or a later version of Outlook to create or send an e-mail that is protected by IRM.


## See also


#### Concepts


[SharingItem Object](sharingitem-object-outlook.md)

