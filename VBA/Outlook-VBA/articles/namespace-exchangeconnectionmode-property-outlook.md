---
title: NameSpace.ExchangeConnectionMode Property (Outlook)
keywords: vbaol11.chm776
f1_keywords:
- vbaol11.chm776
ms.prod: outlook
api_name:
- Outlook.NameSpace.ExchangeConnectionMode
ms.assetid: 4b9f7917-5340-cf72-d690-ac5a7b8d4792
ms.date: 06/08/2017
---


# NameSpace.ExchangeConnectionMode Property (Outlook)

Returns an  **[OlExchangeConnectionMode](olexchangeconnectionmode-enumeration-outlook.md)** constant that indicates the connection mode of the user's primary Exchange account. Read-only.


## Syntax

 _expression_ . **ExchangeConnectionMode**

 _expression_ A variable that represents a **[NameSpace](namespace-object-outlook.md)** object.


## Remarks

If the  **ExchangeConnectionMode** property is **olOffline** or **olDisconnected** , the **[NameSpace.Offline](namespace-offline-property-outlook.md)** property returns **True** . If the **ExchangeConnectionMode** property is **olOnline** , **olConnected** , or **olConnectedHeaders** , the **NameSpace.Offline** property returns **False** .


## Example

The following Microsoft Visual Basic for Applications (VBA) example marks the items that are sent with high importance for download if the connection mode is 'Connected Headers' and the download state is 'Header Only' in the  **Inbox** folder.


```vb
Sub MarkHighImportance() 
 
 Dim myNamespace As Outlook.NameSpace 
 
 Dim mpfInbox As Outlook.Folder 
 
 Dim obj As Object 
 
 Dim ctr As Integer 
 
 Dim i As Integer 
 
 
 
 Set myNamespace = Application.GetNamespace("MAPI") 
 
 Set mpfInbox = myNamespace.GetDefaultFolder(olFolderInbox) 
 
 ctr = mpfInbox.Items.count 
 
 If (myNamespace.ExchangeConnectionMode = olConnectedHeaders) Then 
 
 For i = 1 To ctr 
 
 Set obj = mpfInbox.Items.Item(i) 
 
 If (obj.Importance <> olImportanceHigh And obj.DownloadState = olHeaderOnly) Then 
 
 obj.MarkForDownload = olMarkedForDownload 
 
 End If 
 
 Next 
 
 End If 
 
End Sub
```


## See also


#### Concepts


[NameSpace Object](namespace-object-outlook.md)

