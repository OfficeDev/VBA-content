---
title: PostItem.MarkForDownload Property (Outlook)
keywords: vbaol11.chm1562
f1_keywords:
- vbaol11.chm1562
ms.prod: outlook
api_name:
- Outlook.PostItem.MarkForDownload
ms.assetid: 1b263e7f-923b-89ba-a63c-18af6efa7d38
ms.date: 06/08/2017
---


# PostItem.MarkForDownload Property (Outlook)

Returns or sets an  **[OlRemoteStatus](olremotestatus-enumeration-outlook.md)** constant that determines the status of an item once it is received by a remote user. Read/write.


## Syntax

 _expression_ . **MarkForDownload**

 _expression_ A variable that represents a **PostItem** object.


## Remarks

This property gives remote users with less-than-ideal data-transfer capabilities increased messaging flexibility.


## Example

The following example searches through the user's  **Inbox** for items that have not yet been fully downloaded. If any items are found that are not fully downloaded, a message is displayed and the item is marked for download.


```vb
Sub DownloadItems() 
 
 Dim mpfInbox As Outlook.Folder 
 
 Dim obj As Object 
 
 Dim i As Integer 
 
 
 
 Set mpfInbox = Application.GetNamespace("MAPI").GetDefaultFolder(olFolderInbox) 
 
 'Loop all items in the Inbox folder 
 
 For i = 1 To mpfInbox.Items.Count 
 
 Set obj = mpfInbox.Items.Item(i) 
 
 'Verify if the state of the item is olHeaderOnly 
 
 If obj.DownloadState = olHeaderOnly Then 
 
 MsgBox ("This item has not been fully downloaded.") 
 
 'Mark the item to be downloaded. 
 
 obj.MarkForDownload = olMarkedForDownload 
 
 End If 
 
 Next 
 
End Sub
```


## See also


#### Concepts


[PostItem Object](postitem-object-outlook.md)

