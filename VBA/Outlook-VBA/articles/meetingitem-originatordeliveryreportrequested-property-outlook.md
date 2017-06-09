---
title: MeetingItem.OriginatorDeliveryReportRequested Property (Outlook)
keywords: vbaol11.chm1443
f1_keywords:
- vbaol11.chm1443
ms.prod: outlook
api_name:
- Outlook.MeetingItem.OriginatorDeliveryReportRequested
ms.assetid: 7dfa8dfe-0268-57d8-0ba2-7f69789d4ce9
ms.date: 06/08/2017
---


# MeetingItem.OriginatorDeliveryReportRequested Property (Outlook)

Returns or sets a  **Boolean** value that determines whether the originator of the meeting item or mail message will receive a delivery report. Read/write.


## Syntax

 _expression_ . **OriginatorDeliveryReportRequested**

 _expression_ A variable that represents a **MeetingItem** object.


## Remarks

Each transport provider that handles your message sends you a single delivery notification containing the names and addresses of each recipient to whom it was delivered. Delivery does not imply that the message has been read.  **True** if the originator requested a delivery receipt on the message.

The  **OriginatorDeliveryReportRequested** property corresponds to the MAPI property **PidTagOriginatorDeliveryReportRequested** .


## See also


#### Concepts


[MeetingItem Object](meetingitem-object-outlook.md)

