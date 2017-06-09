---
title: ReportItem.ReplyAll Event (Outlook)
ms.prod: outlook
api_name:
- Outlook.ReportItem.ReplyAll
ms.assetid: b5724798-8c73-13ce-23d4-9d7ec8147f44
ms.date: 06/08/2017
---


# ReportItem.ReplyAll Event (Outlook)

Occurs when the user selects the  **ReplyAll** action for an item (which is an instance of the parent object).


## Syntax

 _expression_ . **ReplyAll**( **_Response_** , **_Cancel_** )

 _expression_ A variable that represents a **ReportItem** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Response_|Required| **Object**|The new item being sent in response to the original message.|
| _Cancel_|Required| **Boolean**| **False** when the event occurs. If the event procedure sets this argument to **True** , the reply all operation is not completed and the new item is not displayed.|

## Remarks

Returns the reply as a  **[MailItem](mailitem-object-outlook.md)** object.


## See also


#### Concepts


[ReportItem Object](reportitem-object-outlook.md)

