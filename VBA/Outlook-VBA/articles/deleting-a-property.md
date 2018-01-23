---
title: Deleting a Property
ms.prod: outlook
ms.assetid: 69d97b27-f60e-6c7a-36c8-a10986101219
ms.date: 06/08/2017
---


# Deleting a Property

Outlook provides several ways to remove a custom property.

|_ObjectProperty_|[UserProperties.Remove](userproperties-remove-method-outlook.md)|[ItemProperties.Remove](itemproperties-remove-method-outlook.md)|[PropertyAccessor.DeleteProperty](propertyaccessor-deleteproperty-method-outlook.md)|[PropertyAccessor.DeleteProperties](propertyaccessor-deleteproperties-method-outlook.md)|
|:-----|:-----|:-----|:-----|:-----|
|**Action**|Removes a custom property specified by _Index_ in the **[UserProperties](userproperties-object-outlook.md)** collection for the item. The **UserProperties** collection is one-based.|Removes a custom property specified by _Index_ in the **[ItemProperties](itemproperties-object-outlook.md)** collection for the item. The **ItemProperties** collection is zero-based. You can only remove custom properties in the collection, and they are denoted by **[IsUserProperty](itemproperty-isuserproperty-property-outlook.md)**. You cannot remove explicit built-in properties.|Removes a custom property specified by _SchemaName_, provided that the property is not read-only and the caller has permission to delete the property (for example, the caller is the owner of the folder to which the property has been added). You cannot remove a built-in Outlook or MAPI property.|For each custom property in _SchemaNames_, removes it provided that the same conditions described in the **PropertyAccessor.DeleteProperty** column are true. Any error will be returned in the corresponding element in the resultant error array.|
|**Applicable objects**|All [Outlook item objects](outlook-item-objects.md) except Office document items (**[DocumentItem](documentitem-object-outlook.md)** objects).|All Outlook item objects except Office document items (**DocumentItem** objects).|All Outlook item objects excluding the **DocumentItem** object, and any of the following objects: **[AddressEntry](addressentry-object-outlook.md)**, **[AddressList](addresslist-object-outlook.md)**, **[Attachment](attachment-object-outlook.md)**, **[ExchangeDistributionList](exchangedistributionlist-object-outlook.md)**, **[ExchangeUser](exchangeuser-object-outlook.md)**, **[Folder](folder-object-outlook.md)**, **[Recipient](recipient-object-outlook.md)**, and **[Store](store-object-outlook.md)** objects.|Same objects as listed in the **DeleteProperty** column.|




