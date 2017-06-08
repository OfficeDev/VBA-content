---
title: TableView.ShowFullConversations Property (Outlook)
keywords: vbaol11.chm3516
f1_keywords:
- vbaol11.chm3516
ms.prod: outlook
api_name:
- Outlook.TableView.ShowFullConversations
ms.assetid: 126cab84-5276-43bd-c19c-2d442e5a2aad
ms.date: 06/08/2017
---


# TableView.ShowFullConversations Property (Outlook)

Returns or sets a  **Boolean** value that indicates whether to display conversation items from other folders, such as the Sent Items folder, as part of the conversation in the table view. Read/write.


## Syntax

 _expression_ . **ShowFullConversations**

 _expression_ A variable that represents a **[TableView](tableview-object-outlook.md)** object.


## Remarks

The  **ShowFullConversations** property takes effect only if the current table view displays items by date and conversation.

 **True** indicates that items of the same conversation are displayed by default as part of the conversation in the table view. **False** indicates that only conversation items in the current folder are displayed in the table view.

The  **ShowFullConversations** property is analogous to selecting **Show Messages from Other Folders** in the **Conversations** menu of the **Arrangement** group on the **View** tab of the ribbon.


## See also


#### Concepts


[TableView Object](tableview-object-outlook.md)

