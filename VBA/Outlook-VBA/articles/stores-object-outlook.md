---
title: Stores Object (Outlook)
keywords: vbaol11.chm3019
f1_keywords:
- vbaol11.chm3019
ms.prod: outlook
api_name:
- Outlook.Stores
ms.assetid: 8915a8e4-9c22-21d5-c492-051d393ce5f7
ms.date: 06/08/2017
---


# Stores Object (Outlook)

A set of  **[Store](store-object-outlook.md)** objects representing all the stores available in the current profile.


## Remarks

You can use the  **Stores** and **Store** objects to enumerate all folders and search folders on all stores in the current session. For more information on storing Outlook items in folders and stores, see[Storing Outlook Items](http://msdn.microsoft.com/library/e4a639a4-10b2-7665-9261-19d6e7707e48%28Office.15%29.aspx).


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


## Events



|**Name**|
|:-----|
|[BeforeStoreRemove](http://msdn.microsoft.com/library/b21d4854-3da5-5c01-cbc1-098bb505466e%28Office.15%29.aspx)|
|[StoreAdd](http://msdn.microsoft.com/library/26e7eddc-9c5a-ffff-d574-afa48e5953d8%28Office.15%29.aspx)|

## Methods



|**Name**|
|:-----|
|[Item](http://msdn.microsoft.com/library/b516241a-7baf-b04b-027d-25de80058fbe%28Office.15%29.aspx)|

## Properties



|**Name**|
|:-----|
|[Application](http://msdn.microsoft.com/library/9605ade2-fe86-30a6-ea1d-787498bf20a5%28Office.15%29.aspx)|
|[Class](http://msdn.microsoft.com/library/fb2b9b17-052c-9b25-53ee-b8fcd9e72cc8%28Office.15%29.aspx)|
|[Count](http://msdn.microsoft.com/library/218d55b5-8394-146b-46eb-d57f444688e8%28Office.15%29.aspx)|
|[Parent](http://msdn.microsoft.com/library/d737cf58-fc6e-a6a1-5144-c294ffbcc314%28Office.15%29.aspx)|
|[Session](http://msdn.microsoft.com/library/aea9466c-4b22-10fa-7938-d12f4f193148%28Office.15%29.aspx)|

## See also


#### Other resources


[Outlook Object Model Reference](http://msdn.microsoft.com/library/73221b13-d8d8-99b8-3394-b95dbbfd5ddc%28Office.15%29.aspx)
[Stores Object Members](http://msdn.microsoft.com/library/f3fec99a-54b2-c13e-d96a-c8c5e2429f99%28Office.15%29.aspx)
