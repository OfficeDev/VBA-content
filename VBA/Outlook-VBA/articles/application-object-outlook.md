---
title: Application Object (Outlook)
keywords: vbaol11.chm2991
f1_keywords:
- vbaol11.chm2991
ms.prod: outlook
api_name:
- Outlook.Application
ms.assetid: 797003e7-ecd1-eccb-eaaf-32d6ddde8348
ms.date: 06/08/2017
---


# Application Object (Outlook)

Represents the entire Microsoft Outlook application.


## Remarks

 This is the only object in the hierarchy that can be returned by using the **[CreateObject](http://msdn.microsoft.com/library/09b6ff5b-a750-c07d-7499-c1f8a00214fe%28Office.15%29.aspx)** method or the intrinsic Visual Basic **GetObject** function.

The Outlook  **Application** object has several purposes:


- As the root object, it allows access to other objects in the Outlook hierarchy.
    
- It allows direct access to a new item created by using  **[CreateItem](http://msdn.microsoft.com/library/e5fbf367-db16-5042-823e-68e6b805e612%28Office.15%29.aspx)**, without having to traverse the object hierarchy.
    
- It allows access to the active interface objects (the explorer and the inspector).
    
When you use Automation to control Outlook from another application, you use the  **CreateObject** method to create an Outlook **Application** object.


## Example

The following Visual Basic for Applications (VBA) example starts Outlook (if it's not already running) and opens the default Inbox folder.


```
Set myNameSpace = Application.GetNameSpace("MAPI") 
 
Set myFolder= _ 
 
 myNameSpace.GetDefaultFolder(olFolderInbox) 
 
myFolder.Display
```

The following Visual Basic for Applications (VBA) example uses the  **Application** object to create and open a new contact.




```
Set myItem = Application.CreateItem(olContactItem) 
 
myItem.Display
```


## Events



|**Name**|
|:-----|
|[AdvancedSearchComplete](http://msdn.microsoft.com/library/4f33ad44-20a3-62cd-aa1b-db74581ebb3c%28Office.15%29.aspx)|
|[AdvancedSearchStopped](http://msdn.microsoft.com/library/a1a4ec9f-c0e3-6acd-b63c-89194ed70efd%28Office.15%29.aspx)|
|[BeforeFolderSharingDialog](http://msdn.microsoft.com/library/e06257eb-f2d9-63cf-1220-dda55ee0ea14%28Office.15%29.aspx)|
|[ItemLoad](http://msdn.microsoft.com/library/aed0656d-4e5a-550a-1116-76773215a897%28Office.15%29.aspx)|
|[ItemSend](http://msdn.microsoft.com/library/54f506ea-87a2-29b9-2b33-67bc87167933%28Office.15%29.aspx)|
|[MAPILogonComplete](http://msdn.microsoft.com/library/db6f7cf8-2a45-560f-f592-613de86e08e2%28Office.15%29.aspx)|
|[NewMail](http://msdn.microsoft.com/library/cfc848e8-98b1-163a-c177-53993c20bb14%28Office.15%29.aspx)|
|[NewMailEx](http://msdn.microsoft.com/library/3b6873a3-0ccf-0e46-1cac-0eeabb3a896b%28Office.15%29.aspx)|
|[OptionsPagesAdd](http://msdn.microsoft.com/library/aa13cd97-de96-00f8-a532-ca8ee9b00343%28Office.15%29.aspx)|
|[Quit](http://msdn.microsoft.com/library/ecf0b50b-db6f-7eaf-90bd-bae942bf9287%28Office.15%29.aspx)|
|[Reminder](http://msdn.microsoft.com/library/f8c9fa87-3daa-58e1-7b8d-3c819cd4cab2%28Office.15%29.aspx)|
|[Startup](http://msdn.microsoft.com/library/d4724d96-2572-b1e3-e202-0bfffb5cf7d5%28Office.15%29.aspx)|

## Methods



|**Name**|
|:-----|
|[ActiveExplorer](http://msdn.microsoft.com/library/f6dd27c0-4319-c7fc-191f-8b3b2ea319d3%28Office.15%29.aspx)|
|[ActiveInspector](http://msdn.microsoft.com/library/3f2b6491-7b4b-8165-327e-b319711d5656%28Office.15%29.aspx)|
|[ActiveWindow](http://msdn.microsoft.com/library/5f5b4e8b-61e4-417b-6b0c-14d1ccb41594%28Office.15%29.aspx)|
|[AdvancedSearch](http://msdn.microsoft.com/library/7b433d8b-08b9-dff1-b854-287d76b47a90%28Office.15%29.aspx)|
|[CopyFile](http://msdn.microsoft.com/library/dc848d48-23e0-d0a9-049d-b2ae414151d5%28Office.15%29.aspx)|
|[CreateItem](http://msdn.microsoft.com/library/e5fbf367-db16-5042-823e-68e6b805e612%28Office.15%29.aspx)|
|[CreateItemFromTemplate](http://msdn.microsoft.com/library/5e6c0ec4-779d-3743-afdb-606ad512ba95%28Office.15%29.aspx)|
|[CreateObject](http://msdn.microsoft.com/library/09b6ff5b-a750-c07d-7499-c1f8a00214fe%28Office.15%29.aspx)|
|[GetNamespace](http://msdn.microsoft.com/library/6175d0d9-5a61-ce45-35c0-b70895d757b3%28Office.15%29.aspx)|
|[GetObjectReference](http://msdn.microsoft.com/library/426ade68-155b-9076-b3f8-4108f44688b0%28Office.15%29.aspx)|
|[IsSearchSynchronous](http://msdn.microsoft.com/library/cd757b43-5e3f-1504-9944-7431bda6f004%28Office.15%29.aspx)|
|[Quit](http://msdn.microsoft.com/library/664bc8ba-ad97-8d4f-02f9-7f9bdd04beea%28Office.15%29.aspx)|
|[RefreshFormRegionDefinition](http://msdn.microsoft.com/library/35183f18-7c59-80c5-e281-af15afe39198%28Office.15%29.aspx)|

## Properties



|**Name**|
|:-----|
|[Application](http://msdn.microsoft.com/library/c49cfea1-d126-75eb-fb3d-6f040526cef0%28Office.15%29.aspx)|
|[Assistance](http://msdn.microsoft.com/library/14d6eb82-82ab-ea67-6a0b-103a535b8d41%28Office.15%29.aspx)|
|[Class](http://msdn.microsoft.com/library/5bfb1d90-8c16-fdbe-374f-0b10d64915c3%28Office.15%29.aspx)|
|[COMAddIns](http://msdn.microsoft.com/library/f911199d-dc2e-9b88-d807-a5737a39f29e%28Office.15%29.aspx)|
|[DefaultProfileName](http://msdn.microsoft.com/library/53c6a189-9337-6413-72e5-bf6ea8794361%28Office.15%29.aspx)|
|[Explorers](http://msdn.microsoft.com/library/bbbdbd6e-a238-8108-fbbd-5f7d7821aaa7%28Office.15%29.aspx)|
|[Inspectors](http://msdn.microsoft.com/library/c2dde847-d033-90e3-30d2-62ff375d6843%28Office.15%29.aspx)|
|[IsTrusted](http://msdn.microsoft.com/library/4caeb41a-9cc3-1195-22a9-ad8eae12ce53%28Office.15%29.aspx)|
|[LanguageSettings](http://msdn.microsoft.com/library/8367a51a-629f-3349-fe0b-a978b2bbc9a5%28Office.15%29.aspx)|
|[Name](http://msdn.microsoft.com/library/a0ac022e-4d46-fffb-aa13-f95249e30bdb%28Office.15%29.aspx)|
|[Parent](http://msdn.microsoft.com/library/d83e85a0-f3d4-bf95-0568-0411a5d09350%28Office.15%29.aspx)|
|[PickerDialog](http://msdn.microsoft.com/library/14acc98b-c234-d59b-d089-d6782ffb08a0%28Office.15%29.aspx)|
|[ProductCode](http://msdn.microsoft.com/library/cdb4678a-fa6b-7d4f-b0b1-b34811749bf5%28Office.15%29.aspx)|
|[Reminders](http://msdn.microsoft.com/library/1f5428f0-6362-a691-2fad-c80e48dce3f5%28Office.15%29.aspx)|
|[Session](http://msdn.microsoft.com/library/720b2849-fe01-afb3-363c-f3bf0cd7d872%28Office.15%29.aspx)|
|[TimeZones](http://msdn.microsoft.com/library/920e55d1-9914-fa74-101a-921083328d23%28Office.15%29.aspx)|
|[Version](http://msdn.microsoft.com/library/08a74ab8-7e02-3956-1827-4b6690acdec1%28Office.15%29.aspx)|

## See also


#### Other resources

[Outlook Object Model Reference](http://msdn.microsoft.com/library/73221b13-d8d8-99b8-3394-b95dbbfd5ddc%28Office.15%29.aspx)

