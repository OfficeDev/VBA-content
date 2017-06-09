---
title: Folder Object (Outlook)
keywords: vbaol11.chm3020
f1_keywords:
- vbaol11.chm3020
ms.prod: outlook
api_name:
- Outlook.Folder
ms.assetid: 3cf6cda8-6d70-666e-2643-9d9c5b9cacfc
ms.date: 06/08/2017
---


# Folder Object (Outlook)

Represents an Outlook folder.


## Remarks

A  **Folder** object can contain other **Folder** objects, as well as Outlook items. Use the **Folders** property of a **[NameSpace](namespace-object-outlook.md)** object or another **Folder** object to return the set of folders in a **NameSpace** or under a folder. You can navigate nested folders by starting from a top-level folder, say the Inbox, and using a combination of the **[Folder.Folders](folder-folders-property-outlook.md)** property, which returns the set of folders underneath a **Folder** object in the hierarchy, and the **[Folders.Item](http://msdn.microsoft.com/library/96a462c2-fa55-62dc-48a4-6464966b84ce%28Office.15%29.aspx)** method, which returns a folder within the **[Folders](http://msdn.microsoft.com/library/0c814c3c-74fc-414c-982d-a0097fcb35c2%28Office.15%29.aspx)** collection.

There is a set of folders within an Outlook data store that supports the default functionality of Outlook. Use  **[NameSpace.GetDefaultFolder](http://msdn.microsoft.com/library/761b8b53-dd4d-43e4-c8f0-69cefdf0c77a%28Office.15%29.aspx)**, specifying an _index_ that is one of the constants in the **[OlDefaultFolders](http://msdn.microsoft.com/library/1a17abd8-09b9-d3e1-2d93-0a4d5580a950%28Office.15%29.aspx)** enumeration to return one of the default Outlook folders in the Outlook **NameSpace** object.

 While generally it is a good practice to place items that serve the same functionality in the same folder, a folder can contain items of different types. For example, by default, the Calendar folder can contain **[AppointmentItem](appointmentitem-object-outlook.md)** and **[MeetingItem](meetingitem-object-outlook.md)** objects, and the Contacts folder can contain **[ContactItem](contactitem-object-outlook.md)** and **[DistListItem](distlistitem-object-outlook.md)** objects. In general, when enumerating items in a folder, do not assume the type of an item in the folder; check the message class of the item before accessing properties that are applicable to the item.

 Use the **[Folders.Add](http://msdn.microsoft.com/library/20ced7ad-779c-a9b0-267e-6d729c0eb822%28Office.15%29.aspx)** method to add a folder to the **Folders** object. The **Add** method has an optional argument that can be used to specify the type of items that can be stored in that folder. By default, folders created inside another folder inherit the type of the parent folder.

 Note that when items of a specific type are saved, they are saved directly into their corresponding default folder. For example, when the **[MeetingItem.GetAssociatedAppointment](http://msdn.microsoft.com/library/8344d40d-5c1d-ead3-87cb-fd795b831712%28Office.15%29.aspx)** method is applied to a **MeetingItem** in the Inbox folder, the appointment that is returned will be saved to the default Calendar folder.


## Events



|**Name**|
|:-----|
|[BeforeFolderMove](http://msdn.microsoft.com/library/c085f0cf-3d91-db84-aab9-18c7b46a04d2%28Office.15%29.aspx)|
|[BeforeItemMove](http://msdn.microsoft.com/library/db75bc05-c80e-e6b8-d017-2150bc942712%28Office.15%29.aspx)|

## Methods



|**Name**|
|:-----|
|[AddToPFFavorites](http://msdn.microsoft.com/library/d3926957-bf6d-ad4d-9c24-bfc5037ba9fd%28Office.15%29.aspx)|
|[CopyTo](http://msdn.microsoft.com/library/ddd010e2-54af-f291-b9a9-92cc55a83cca%28Office.15%29.aspx)|
|[Delete](http://msdn.microsoft.com/library/3df0f063-3f41-e3b7-d1e3-7ea08970c56d%28Office.15%29.aspx)|
|[Display](http://msdn.microsoft.com/library/cde389e0-5ec9-8261-5ec0-9a5ba4f8776d%28Office.15%29.aspx)|
|[GetCalendarExporter](http://msdn.microsoft.com/library/7c67e208-65dd-8904-4b6f-8ec2df4e530d%28Office.15%29.aspx)|
|[GetCustomIcon](http://msdn.microsoft.com/library/49a3da64-2b2f-76db-0053-88e35141cca0%28Office.15%29.aspx)|
|[GetExplorer](http://msdn.microsoft.com/library/f60bf373-802e-cb93-2152-bc6c8945edb1%28Office.15%29.aspx)|
|[GetStorage](http://msdn.microsoft.com/library/cc5ee63b-7d11-6340-8392-8b35a689a28c%28Office.15%29.aspx)|
|[GetTable](http://msdn.microsoft.com/library/08d184cb-0c41-01b1-abc5-305476380f8b%28Office.15%29.aspx)|
|[MoveTo](http://msdn.microsoft.com/library/5e8ece38-aaba-4971-643e-969956c2a196%28Office.15%29.aspx)|
|[SetCustomIcon](http://msdn.microsoft.com/library/d368547b-e90c-85ec-7d5c-e48cbe8eb42e%28Office.15%29.aspx)|

## Properties



|**Name**|
|:-----|
|[AddressBookName](http://msdn.microsoft.com/library/e80535e9-216f-03a6-36a1-3776b5862e96%28Office.15%29.aspx)|
|[Application](http://msdn.microsoft.com/library/525cac55-6eb7-a7c5-8949-a17cf6e6bc33%28Office.15%29.aspx)|
|[Class](http://msdn.microsoft.com/library/6ec62401-52b2-acb4-af3f-b160ea5e28fc%28Office.15%29.aspx)|
|[CurrentView](http://msdn.microsoft.com/library/42af4345-60f1-10cd-66e5-517ca002284b%28Office.15%29.aspx)|
|[CustomViewsOnly](http://msdn.microsoft.com/library/b94b60f3-acd8-a968-f333-cb6d4bae8683%28Office.15%29.aspx)|
|[DefaultItemType](http://msdn.microsoft.com/library/5a08d9aa-6bb7-0917-6d46-cb27cd03dace%28Office.15%29.aspx)|
|[DefaultMessageClass](http://msdn.microsoft.com/library/f8daf970-4ae1-6713-c7f8-4420d952b823%28Office.15%29.aspx)|
|[Description](http://msdn.microsoft.com/library/e8cddfad-b071-b641-268b-dc10cfed20f8%28Office.15%29.aspx)|
|[EntryID](http://msdn.microsoft.com/library/338ade5a-b267-8bc2-35b7-221c071506aa%28Office.15%29.aspx)|
|[FolderPath](http://msdn.microsoft.com/library/40a588fa-0962-bc01-f8ac-39f0bab2092c%28Office.15%29.aspx)|
|[Folders](folder-folders-property-outlook.md)|
|[InAppFolderSyncObject](http://msdn.microsoft.com/library/d9e94fb7-add5-65d5-d2bc-e23bdfa11078%28Office.15%29.aspx)|
|[IsSharePointFolder](http://msdn.microsoft.com/library/fc2e2645-d6e0-0bc0-29a2-8cc17f456225%28Office.15%29.aspx)|
|[Items](http://msdn.microsoft.com/library/441820e7-5fe8-e5ef-83c0-9c87fd3dc9e3%28Office.15%29.aspx)|
|[Name](http://msdn.microsoft.com/library/ec03a345-8c06-f234-e1e9-ecdc54495ed2%28Office.15%29.aspx)|
|[Parent](http://msdn.microsoft.com/library/0671c1d3-c25e-b9c7-3c07-bd83c9f01ae4%28Office.15%29.aspx)|
|[PropertyAccessor](http://msdn.microsoft.com/library/8b6fb7a7-a87d-3df3-ae74-19447bc31a0e%28Office.15%29.aspx)|
|[Session](http://msdn.microsoft.com/library/b446d857-4f41-085f-7303-3e5050e33bba%28Office.15%29.aspx)|
|[ShowAsOutlookAB](http://msdn.microsoft.com/library/bb74591b-a3ea-efbd-e7b2-f374f1974be8%28Office.15%29.aspx)|
|[ShowItemCount](http://msdn.microsoft.com/library/3ce32c47-5f92-82ca-5ac3-b3d6f24e5f36%28Office.15%29.aspx)|
|[Store](http://msdn.microsoft.com/library/347d3031-01cf-a248-4abc-f749feb811a4%28Office.15%29.aspx)|
|[StoreID](http://msdn.microsoft.com/library/8b2657b7-0c69-d8ad-147b-482303ebd10f%28Office.15%29.aspx)|
|[UnReadItemCount](http://msdn.microsoft.com/library/f669af8e-8d4a-613b-1d82-6a3be8dc67cd%28Office.15%29.aspx)|
|[UserDefinedProperties](http://msdn.microsoft.com/library/4293bcb8-855e-4c6d-9718-ba8c5862b3bd%28Office.15%29.aspx)|
|[Views](http://msdn.microsoft.com/library/24ef613a-9832-032c-4e68-1001a0385b11%28Office.15%29.aspx)|
|[WebViewOn](http://msdn.microsoft.com/library/9b483d0e-dea0-9b3e-8ce9-fc136857a428%28Office.15%29.aspx)|
|[WebViewURL](http://msdn.microsoft.com/library/6a7630c2-5c16-f63f-a435-987c7c1251b8%28Office.15%29.aspx)|

## See also


#### Other resources


[Outlook Object Model Reference](http://msdn.microsoft.com/library/73221b13-d8d8-99b8-3394-b95dbbfd5ddc%28Office.15%29.aspx)
[Folder Object Members](http://msdn.microsoft.com/library/788acd42-377a-1803-7713-50e45086e2d1%28Office.15%29.aspx)
