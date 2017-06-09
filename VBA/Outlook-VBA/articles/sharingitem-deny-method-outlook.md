---
title: SharingItem.Deny Method (Outlook)
keywords: vbaol11.chm692
f1_keywords:
- vbaol11.chm692
ms.prod: outlook
api_name:
- Outlook.SharingItem.Deny
ms.assetid: f2a5af98-280d-48f1-f6c3-3d17a2654774
ms.date: 06/08/2017
---


# SharingItem.Deny Method (Outlook)

Denies a sharing request and sends a sharing response to the sender of the  **[SharingItem](sharingitem-object-outlook.md)** .


## Syntax

 _expression_ . **Deny**

 _expression_ An expression that returns a **SharingItem** object.


### Return Value

A  **SharingItem** object that represents the sharing response.


## Remarks

The  **Deny** method can only be called on **SharingItem** objects with a **[Type](sharingitem-type-property-outlook.md)** property value of **olSharingMsgTypeRequest** or **olSharingMsgTypeInviteAndRequest** .

This method generates a new  **SharingItem** object and sets the **Type** property of the new object to **olSharingMsgTypeResponseDeny** . The **SharingItem** is not immediately sent to the sender of the sharing request, however, so you can edit the sharing response as needed.


 **Note**  Sharing is denied immediately after this method is called, regardless of whether the sharing response was received.


## See also


#### Concepts


[SharingItem Object](sharingitem-object-outlook.md)

