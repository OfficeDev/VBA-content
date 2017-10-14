---
title: Items Object (Outlook)
keywords: vbaol11.chm2998
f1_keywords:
- vbaol11.chm2998
ms.prod: outlook
api_name:
- Outlook.Items
ms.assetid: 3a99730b-e62a-5ca6-f6ec-911c95173242
ms.date: 06/08/2017
---


# Items Object (Outlook)

Contains a collection of [Outlook item objects](http://msdn.microsoft.com/library/6ea4babf-facf-4018-ef5a-4a484e55153a%28Office.15%29.aspx) in a folder.


## Remarks

Use the  **[Items](http://msdn.microsoft.com/library/441820e7-5fe8-e5ef-83c0-9c87fd3dc9e3%28Office.15%29.aspx)** property to return the **Items** object of a **[Folder](folder-object-outlook.md)** object.

Use  **Items** ( _index_ ), where _index_ is the name or index number, to return a single Outlook item.


 **Note**  The index for the  **Items** collection starts at 1, and the items in the **Items** collection object are not guaranteed to be in any particular order.


## Example

The following Microsoft Visual Basic for Applications (VBA) example returns the first item in the  **Inbox** with the Subject "Need your advice."






```
Sub GetItem() 
 
 Dim myNameSpace As Outlook.NameSpace 
 
 Dim myFolder As Outlook.Folder 
 
 Dim myItem As Object 
 
 
 
 Set myNameSpace = Application.GetNameSpace("MAPI") 
 
 Set myFolder = _ 
 
 myNameSpace.GetDefaultFolder(olFolderInbox) 
 
 Set myItem = myFolder.Items("Need your advice") 
 
 myItem.Display 
 
End sub
```

The following VBA example returns the first item in the  **Inbox**. In Microsoft Office Outlook 2003 or later, the  **Items** object returns the items in an Offline Folders file (.ost) in the reverse order.






```
Sub GetItem() 
 
 Dim myNameSpace As Outlook.NameSpace 
 
 Dim myFolder As Outlook.Folder 
 
 Dim myItem As Object 
 
 
 
 Set myNameSpace = Application.GetNameSpace("MAPI") 
 
 Set myFolder = _ 
 
 myNameSpace.GetDefaultFolder(olFolderInbox) 
 
 Set myItem = myFolder.Items(1) 
 
 myItem.Display 
 
End sub
```


## Events



|**Name**|
|:-----|
|[ItemAdd](http://msdn.microsoft.com/library/e46f5958-aff8-3a6b-b3df-5c4352b6c3d9%28Office.15%29.aspx)|
|[ItemChange](http://msdn.microsoft.com/library/6478357e-2a5a-300a-24e6-c125f8c81edd%28Office.15%29.aspx)|
|[ItemRemove](http://msdn.microsoft.com/library/c1b2d9cd-ab32-2c4a-85fa-9412c190ac4f%28Office.15%29.aspx)|

## Methods



|**Name**|
|:-----|
|[Add](http://msdn.microsoft.com/library/0ee68068-1452-0f29-b85a-88b801ac0448%28Office.15%29.aspx)|
|[Find](http://msdn.microsoft.com/library/e7a791d8-b80b-df07-84a3-a85acabfcf80%28Office.15%29.aspx)|
|[FindNext](http://msdn.microsoft.com/library/2530f640-e024-3567-f539-6bdbf645401d%28Office.15%29.aspx)|
|[GetFirst](http://msdn.microsoft.com/library/142a6174-118e-6256-0511-8ae9e142e555%28Office.15%29.aspx)|
|[GetLast](http://msdn.microsoft.com/library/d02a20be-19fc-fb6e-feff-b66ca0273beb%28Office.15%29.aspx)|
|[GetNext](http://msdn.microsoft.com/library/01c49c21-d9f9-37c4-8c64-ff8e2b1f9462%28Office.15%29.aspx)|
|[GetPrevious](http://msdn.microsoft.com/library/5dde47f8-2bd8-fdbe-d6e7-b1381e8a97a6%28Office.15%29.aspx)|
|[Item](http://msdn.microsoft.com/library/89a031e0-c0a3-fc22-f485-189df8db45f4%28Office.15%29.aspx)|
|[Remove](http://msdn.microsoft.com/library/d2838c82-d0ac-82cc-eed0-c34d55c67d63%28Office.15%29.aspx)|
|[ResetColumns](http://msdn.microsoft.com/library/0543dd17-1e65-5484-ab21-d4791b3b1194%28Office.15%29.aspx)|
|[Restrict](http://msdn.microsoft.com/library/e3b0cda1-e43d-cc5e-2942-0f54935d9dab%28Office.15%29.aspx)|
|[SetColumns](http://msdn.microsoft.com/library/90206a68-baf8-282c-5793-fee029fed452%28Office.15%29.aspx)|
|[Sort](http://msdn.microsoft.com/library/7cb248a2-6885-8be5-df7b-fd5683081e01%28Office.15%29.aspx)|

## Properties



|**Name**|
|:-----|
|[Application](http://msdn.microsoft.com/library/b55a6901-fbd4-36a1-47e7-2c1e37e0a31c%28Office.15%29.aspx)|
|[Class](http://msdn.microsoft.com/library/783ed46a-fd40-c848-b440-8ea3c5d0e6b9%28Office.15%29.aspx)|
|[Count](http://msdn.microsoft.com/library/c18b06be-3a21-3350-6d14-57c822a85d42%28Office.15%29.aspx)|
|[IncludeRecurrences](http://msdn.microsoft.com/library/7d192112-889c-56ce-aab2-107d751c80c4%28Office.15%29.aspx)|
|[Parent](http://msdn.microsoft.com/library/8e99782a-5654-ae1d-c6d8-9dbfcbcf44ac%28Office.15%29.aspx)|
|[Session](http://msdn.microsoft.com/library/5c385dfc-042e-7649-0f32-5d34e53fca57%28Office.15%29.aspx)|

## See also


#### Other resources


[Items Object Members](http://msdn.microsoft.com/library/bcc2cf6c-b6fb-e1a2-1d5c-d7e2bdf6b7dc%28Office.15%29.aspx)
[Outlook Object Model Reference](http://msdn.microsoft.com/library/73221b13-d8d8-99b8-3394-b95dbbfd5ddc%28Office.15%29.aspx)
