---
title: Attachment.Position Property (Outlook)
keywords: vbaol11.chm2370
f1_keywords:
- vbaol11.chm2370
ms.prod: outlook
api_name:
- Outlook.Attachment.Position
ms.assetid: f280b9f5-3484-ad4c-87f8-1caa8631d808
ms.date: 06/08/2017
---


# Attachment.Position Property (Outlook)

Returns or sets a  **Long** indicating the position of the attachment within the body of the item. Read/write.


## Syntax

 _expression_ . **Position**

 _expression_ A variable that represents an **Attachment** object.


## Remarks

The  **Position** property only works on an attachment for an item where the body format is Rich Text (RTF). If the body format is not RTF, the **Position** property is ignored for a set operation and always returns zero (0) for a get operation.

If you set the  **Position** property to 0 for an item where the body format is RTF, the attachment will be hidden in Outlook's user interface. The attachment is not visible in a view, and the user cannot remove the attachment from the body of the item. The attachment can still be accessed through the **[Attachments](attachments-object-outlook.md)** collection on the item.

The only item that allows programmatic setting of the  **BodyFormat** property is a **[MailItem](mailitem-object-outlook.md)** . Other item types such as Appointment, Contact, and Task are RTF by default.


## See also


#### Concepts


[Attachment Object](attachment-object-outlook.md)

