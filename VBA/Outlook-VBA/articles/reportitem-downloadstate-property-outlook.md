---
title: ReportItem.DownloadState Property (Outlook)
keywords: vbaol11.chm1673
f1_keywords:
- vbaol11.chm1673
ms.prod: outlook
api_name:
- Outlook.ReportItem.DownloadState
ms.assetid: e81a4cc1-b94f-b5cb-7224-68d90c075f8b
ms.date: 06/08/2017
---


# ReportItem.DownloadState Property (Outlook)

Returns a constant that belongs to the  **[OlDownloadState](oldownloadstate-enumeration-outlook.md)** enumeration indicating the download state of the item. Read-only.


## Syntax

 _expression_ . **DownloadState**

 _expression_ A variable that represents a **ReportItem** object.


## Example

The following Microsoft Visual Basic for Applications (VBA) example searches through the user's  **Inbox** for items that have not yet been fully downloaded. If any not yet fully downloaded items are found, a message is displayed to the user, and the item is marked for download.


```vb
Sub DownloadItems() 
 
 Dim mpfInbox As Outlook.Folder 
 
 Dim objItems As Outlook.Items 
 
 Dim obj As Object 
 
 Dim i As Integer 
 
 Dim iCount As Integer 
 
 
 
 Set mpfInbox = Application.GetNamespace("MAPI").GetDefaultFolder(olFolderInbox) 
 
 Set objItems = mpfInbox.Items 
 
 iCount = objItems.Count 
 
 'Loop all items in the Inbox folder 
 
 For i = 1 To iCount 
 
 Set obj = objItems.Item(i) 
 
 'Verify if the state of the item is olHeaderOnly 
 
 If obj.DownloadState = olHeaderOnly Then 
 
 MsgBox "This item has not been fully downloaded." 
 
 'Mark the item to be downloaded 
 
 obj.MarkForDownload = olMarkedForDownload 
 
 obj.Save 
 
 End If 
 
 Next 
 
End Sub
```


## See also


#### Concepts


[ReportItem Object](reportitem-object-outlook.md)

