---
title: MailItem.OriginatorDeliveryReportRequested Property (Outlook)
keywords: vbaol11.chm1339
f1_keywords:
- vbaol11.chm1339
ms.prod: outlook
api_name:
- Outlook.MailItem.OriginatorDeliveryReportRequested
ms.assetid: 89042dd2-4ac1-109d-5f9c-9ed3733032b0
ms.date: 06/08/2017
---


# MailItem.OriginatorDeliveryReportRequested Property (Outlook)

Returns or sets a  **Boolean** value that determines whether the originator of the meeting item or mail message will receive a delivery report. Read/write.


## Syntax

 _expression_ . **OriginatorDeliveryReportRequested**

 _expression_ A variable that represents a **MailItem** object.


## Remarks

Each transport provider that handles your message sends you a single delivery notification containing the names and addresses of each recipient to whom it was delivered. Delivery does not imply that the message has been read.  **True** if the originator requested a delivery receipt on the message.

The  **OriginatorDeliveryReportRequested** property corresponds to the MAPI property **PidTagOriginatorDeliveryReportRequested** .


## See also


#### Concepts


[MailItem Object](mailitem-object-outlook.md)

