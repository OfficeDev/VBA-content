---
title: TableView.AutoPreview Property (Outlook)
keywords: vbaol11.chm2536
f1_keywords:
- vbaol11.chm2536
ms.prod: outlook
api_name:
- Outlook.TableView.AutoPreview
ms.assetid: 51d20d34-5a2f-03f6-cfea-2279d286f067
ms.date: 06/08/2017
---


# TableView.AutoPreview Property (Outlook)

Returns or sets an  **[OlAutoPreview](olautopreview-enumeration-outlook.md)** constant that determines how items are automatically previewed by the **[TableView](tableview-object-outlook.md)** object. Read/write.


## Syntax

 _expression_ . **AutoPreview**

 _expression_ A variable that represents a **TableView** object.


## Example

The following Visual Basic for Applications (VBA) example sets the  **AutoPreview** property to **olAutoPreviewUnread** for every **TableView** object associated with the current **[Folder](folder-object-outlook.md)** object.


```vb
Private Sub PreviewUnreadOnly() 
 
 Dim objFolder As Folder 
 
 Dim objView As View 
 
 Dim objTableView As TableView 
 
 
 
 ' Retrieve a Folder object reference 
 
 ' for the current folder 
 
 Set objFolder = Application.ActiveExplorer.CurrentFolder 
 
 
 
 ' Enumerate through the Views collection for the 
 
 ' Folder object. 
 
 For Each objView In objFolder.Views 
 
 ' Check if the view is a table view. 
 
 If objView.ViewType = olTableView Then 
 
 ' Cast the View object to a TableView object. 
 
 Set objTableView = objView 
 
 
 
 ' Set the view so that only unread messages 
 
 ' are automatically previewed. 
 
 objTableView.AutoPreview = olAutoPreviewUnread 
 
 
 
 ' Save the table view. 
 
 objTableView.Save 
 
 End If 
 
 Next 
 
End Sub
```


## See also


#### Concepts


[TableView Object](tableview-object-outlook.md)

