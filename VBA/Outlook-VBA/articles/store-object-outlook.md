---
title: Store Object (Outlook)
keywords: vbaol11.chm3155
f1_keywords:
- vbaol11.chm3155
ms.prod: outlook
api_name:
- Outlook.Store
ms.assetid: 1eb22fe9-8849-7476-5388-2515b48591b9
ms.date: 06/08/2017
---


# Store Object (Outlook)

Represents a file on the local computer or a network drive that stores e-mail messages and other items for an account in the current profile.


## Remarks

A profile defines one or more e-mail accounts, and each e-mail account is associated with a server of a specific type. For an Exchange server, a store can be on the server, in an Exchange Public folder, or in a local Personal Folders File (.pst) or Offline Folder File (.ost). For a POP3, IMAP, or HTTP e-mail server, a store is a .pst file.

You can use the  **[Stores](stores-object-outlook.md)** and **Store** objects to enumerate all folders and search folders on all stores in the current session. Since getting the root folder or search folders in a store requires the store to be open and opening a store imposes an overhead on performance, you can check the **[Store.IsOpen](http://msdn.microsoft.com/library/05e93457-2d17-39ac-404c-c78c76d2ef72%28Office.15%29.aspx)** property before you decide to pursue the operation.

If you use an Exchange server, you can access other explicit built-in  **Store** properties for store characteristics such as **[ExchangeStoreType](http://msdn.microsoft.com/library/ca6002bd-444d-a111-adca-6f8fafc37ea1%28Office.15%29.aspx)**, **[IsCachedExchange](http://msdn.microsoft.com/library/2f3fbd5d-8cf1-5fdd-6074-f4da4216dcd4%28Office.15%29.aspx)**, and **[IsDataFileStore](http://msdn.microsoft.com/library/76dc73b7-1d19-465f-744f-1209211f2496%28Office.15%29.aspx)**. Use the **[PropertyAccessor](propertyaccessor-object-outlook.md)** object returned by **[Store.PropertyAccessor](http://msdn.microsoft.com/library/4c3ccfc9-8f8a-aa2b-f7f5-5945ffe55f31%28Office.15%29.aspx)** to access other store properties that are not exposed in the Outlook object model.

For more information on storing Outlook items in folders and stores, see [Storing Outlook Items](http://msdn.microsoft.com/library/e4a639a4-10b2-7665-9261-19d6e7707e48%28Office.15%29.aspx).


## Example

The following code sample in Microsoft Visual Basic for Applications (VBA) enumerates all folders on all stores for a session:


```
Sub EnumerateFoldersInStores() 
 
 Dim colStores As Outlook.Stores 
 
 Dim oStore As Outlook.Store 
 
 Dim oRoot As Outlook.Folder 
 
 
 
 On Error Resume Next 
 
 Set colStores = Application.Session.Stores 
 
 For Each oStore In colStores 
 
 Set oRoot = oStore.GetRootFolder 
 
 Debug.Print (oRoot.FolderPath) 
 
 EnumerateFolders oRoot 
 
 Next 
 
End Sub 
 
 
 
Private Sub EnumerateFolders(ByVal oFolder As Outlook.Folder) 
 
 Dim folders As Outlook.folders 
 
 Dim Folder As Outlook.Folder 
 
 Dim foldercount As Integer 
 
 
 
 On Error Resume Next 
 
 Set folders = oFolder.folders 
 
 foldercount = folders.Count 
 
 'Check if there are any folders below oFolder 
 
 If foldercount Then 
 
 For Each Folder In folders 
 
 Debug.Print (Folder.FolderPath) 
 
 EnumerateFolders Folder 
 
 Next 
 
 End If 
 
End Sub
```


## Methods



|**Name**|
|:-----|
|[GetDefaultFolder](http://msdn.microsoft.com/library/f3e87528-6de8-dc59-8d27-f19f6b344044%28Office.15%29.aspx)|
|[GetRootFolder](http://msdn.microsoft.com/library/09da4d57-c33d-6946-cc21-7233e89efb10%28Office.15%29.aspx)|
|[GetRules](http://msdn.microsoft.com/library/06048799-e162-68f9-17c2-d80c25e2c55e%28Office.15%29.aspx)|
|[GetSearchFolders](http://msdn.microsoft.com/library/aed6ba0b-5e20-adb9-6f62-d030a0de2e0b%28Office.15%29.aspx)|
|[GetSpecialFolder](http://msdn.microsoft.com/library/8f768a43-1589-5659-76f3-43afa4b745b6%28Office.15%29.aspx)|
|[RefreshQuotaDisplay](http://msdn.microsoft.com/library/131540a9-f803-29a8-82e1-caa7f14298ef%28Office.15%29.aspx)|
|[CreateUnifiedGroup](http://msdn.microsoft.com/library/45f70f08-f198-22a2-79c5-26dc3247e164%28Office.15%29.aspx)|
|[DeleteUnifiedGroup](http://msdn.microsoft.com/library/53c15736-f88a-33ad-2b21-29a2c9c6d402%28Office.15%29.aspx)|

## Properties



|**Name**|
|:-----|
|[Application](http://msdn.microsoft.com/library/97ea6907-8619-3777-d201-2727a59ff59c%28Office.15%29.aspx)|
|[Categories](http://msdn.microsoft.com/library/597678d0-51f6-45d7-a98a-063344bbcff7%28Office.15%29.aspx)|
|[Class](http://msdn.microsoft.com/library/fcc205ac-a1af-d215-e8b9-91cfd2147634%28Office.15%29.aspx)|
|[DisplayName](http://msdn.microsoft.com/library/785ec583-3553-6002-41b6-d0c6d0028b5a%28Office.15%29.aspx)|
|[ExchangeStoreType](http://msdn.microsoft.com/library/ca6002bd-444d-a111-adca-6f8fafc37ea1%28Office.15%29.aspx)|
|[FilePath](http://msdn.microsoft.com/library/3b0ed312-9304-61a6-7152-5693a0e2f0fe%28Office.15%29.aspx)|
|[IsCachedExchange](http://msdn.microsoft.com/library/2f3fbd5d-8cf1-5fdd-6074-f4da4216dcd4%28Office.15%29.aspx)|
|[IsConversationEnabled](http://msdn.microsoft.com/library/ce333881-a5f3-2115-0ae4-296d15c4bead%28Office.15%29.aspx)|
|[IsDataFileStore](http://msdn.microsoft.com/library/76dc73b7-1d19-465f-744f-1209211f2496%28Office.15%29.aspx)|
|[IsInstantSearchEnabled](http://msdn.microsoft.com/library/0fba75cc-c506-157b-7dfa-ec438e932f5c%28Office.15%29.aspx)|
|[IsOpen](http://msdn.microsoft.com/library/05e93457-2d17-39ac-404c-c78c76d2ef72%28Office.15%29.aspx)|
|[Parent](http://msdn.microsoft.com/library/93484d08-064e-144f-b1da-12eecceb2d83%28Office.15%29.aspx)|
|[PropertyAccessor](http://msdn.microsoft.com/library/4c3ccfc9-8f8a-aa2b-f7f5-5945ffe55f31%28Office.15%29.aspx)|
|[Session](http://msdn.microsoft.com/library/90dc9dc2-41c5-6448-4f42-98d8e4a6f948%28Office.15%29.aspx)|
|[StoreID](http://msdn.microsoft.com/library/fce5fa3a-87dc-68c5-ba5f-ee1430584b5d%28Office.15%29.aspx)|

## See also


#### Other resources


[Outlook Object Model Reference](http://msdn.microsoft.com/library/73221b13-d8d8-99b8-3394-b95dbbfd5ddc%28Office.15%29.aspx)
[Store Object Members](http://msdn.microsoft.com/library/84c1d423-e507-0b3b-6570-33829b94be04%28Office.15%29.aspx)
