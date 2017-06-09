---
title: SharingItem.PermissionTemplateGuid Property (Outlook)
keywords: vbaol11.chm3511
f1_keywords:
- vbaol11.chm3511
ms.prod: outlook
api_name:
- Outlook.SharingItem.PermissionTemplateGuid
ms.assetid: 166c2975-b6be-d1ca-4aa8-ad7deb42c68d
ms.date: 06/08/2017
---


# SharingItem.PermissionTemplateGuid Property (Outlook)

Returns or sets a  **String** that represents the GUID of the template file to be applied to the **[SharingItem](sharingitem-object-outlook.md)** in order to specify Information Rights Management (IRM) permissions. Read/write.


## Syntax

 _expression_ . **PermissionTemplateGuid**

 _expression_ A variable that represents a **SharingItem** object.


## Remarks

This property complements the IRM properties on a  **SharingItem** object; that is, the **[Permission](sharingitem-permission-property-outlook.md)** property and **[PermissionService](sharingitem-permissionservice-property-outlook.md)** properties.

The  **PermissionTemplateGuid** property should be synchronized with the **Permission** property to accurately reflect the permission status of the **SharingItem** . Setting the **PermissionTemplateGuid** property to a valid GUID should also incur setting the **Permission** property to **OlPermission.olPermissionTemplate** .

An empty string value for the  **PermissionTemplateGuid** property means there is no permission template file specified for the **SharingItem** . This occurs when no IRM has been set up (in which case the **Permission** property is **OlPermission.olUnrestricted** ), or the restriction is not to forward the **SharingItem** (in which case the **Permission** property is **OlPermission.olDoNotForward** ).

If you attempt to set the  **PermissionTemplateGuid** property for a received message (that is, the **[Sent](sharingitem-sent-property-outlook.md)** property of the **SharingItem** is **True** ), Microsoft Outlook returns an error.


## See also


#### Concepts


[SharingItem Object](sharingitem-object-outlook.md)

