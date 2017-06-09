---
title: TableView.ShowConversationByDate Property (Outlook)
keywords: vbaol11.chm3515
f1_keywords:
- vbaol11.chm3515
ms.prod: outlook
api_name:
- Outlook.TableView.ShowConversationByDate
ms.assetid: b568d714-93ce-e4a4-c84c-b0870dd565dd
ms.date: 06/08/2017
---


# TableView.ShowConversationByDate Property (Outlook)

Returns or sets a  **Boolean** value that indicates whether items in a conversation are organized vertically left-aligned and ordered by the received date and time, with the most recent item on top. Read/write.


## Syntax

 _expression_ . **ShowConversationByDate**

 _expression_ A variable that represents a **[TableView](tableview-object-outlook.md)** object.


## Remarks



If the table view is not organized by conversation, setting the  **ShowConversationByDate** property does not reorganize any items in the view. To display items by conversation, in the **Conversations** group of the **View** tab in the ribbon, select **Show as Conversations**.

Setting the  **ShowConversationByDate** property to **True** vertically left-aligns conversation items and orders them by their received date and time, with the most recent item on top. This organization in the conversation view is the same as having cleared the **Use Classic Indented View** setting in the **Conversations Settings** menu in the **Conversations** group of the ribbon.

Setting the  **ShowConversationByDate** property to **False** indents conversation items and orders them by their received date and time, with the earliest item on top. The root of each thread of the conversation is displayed first, followed by items belonging to that thread, each left-indented from the last. This organization in the conversation view is the same as having selected the **Use Classic Indented View** setting in the **Conversations Settings** menu in the **Conversations** group of the ribbon.

To apply a change to the  **ShowConversationByDate** property to the view, call the **[Apply](tableview-apply-method-outlook.md)** method. Conversations are then displayed as collapsed in the conversation view. If you expand a conversation, you will see items in the conversation reorganized and displayed the way you set the **ShowConversationByDate** property.


## Example

The following code sample in Microsoft Visual Basic for Applications (VBA) checks if the current view of the current folder is a table view, assumes items are displayed by conversation, sets the  **ShowConversationByDate** property to true, and calls the **Apply** method to apply the organization to the current view.


```vb
Sub GetConversationViewSettings() 
 
 Dim oCurrentFolder As Outlook.folder 
 
 Dim oView As Outlook.tableView 
 
 Set oCurrentFolder = Application.ActiveExplorer.currentfolder 
 
 If oCurrentFolder.currentView.ViewType = olTableView Then 
 
 Set oView = oCurrentFolder.currentView 
 
 oView.ShowConversationByDate = True 
 
 oView.Apply 
 
 End If 
 
End Sub
```


## See also


#### Concepts


[TableView Object](tableview-object-outlook.md)

