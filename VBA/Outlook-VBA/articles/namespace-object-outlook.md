---
title: NameSpace Object (Outlook)
keywords: vbaol11.chm3000
f1_keywords:
- vbaol11.chm3000
ms.prod: outlook
api_name:
- Outlook.NameSpace
ms.assetid: f0dcaa19-07f5-5d42-a3bf-2e42b7885644
ms.date: 06/08/2017
---


# NameSpace Object (Outlook)

Represents an abstract root object for any data source.


## Remarks

The object itself provides methods for logging in and out, accessing storage objects directly by ID, accessing certain special default folders directly, and accessing data sources owned by other users.

Use  **[GetNameSpace](http://msdn.microsoft.com/library/6175d0d9-5a61-ce45-35c0-b70895d757b3%28Office.15%29.aspx)** ("MAPI") to return the Outlook **NameSpace** object from the **[Application](http://msdn.microsoft.com/library/797003e7-ecd1-eccb-eaaf-32d6ddde8348%28Office.15%29.aspx)** object.

The only data source supported is MAPI, which allows access to all Outlook data stored in the user's mail stores.


## Events



|**Name**|
|:-----|
|[AutoDiscoverComplete](http://msdn.microsoft.com/library/b7cac212-4d38-660e-0caf-48f97035f14a%28Office.15%29.aspx)|
|[OptionsPagesAdd](http://msdn.microsoft.com/library/3f4920bd-ab22-90a7-490a-67122dac6c51%28Office.15%29.aspx)|

## Methods



|**Name**|
|:-----|
|[AddStore](http://msdn.microsoft.com/library/c9390982-2408-fda5-a14d-de6f0daaadf1%28Office.15%29.aspx)|
|[AddStoreEx](http://msdn.microsoft.com/library/15b8948d-cbe4-a499-ec03-b1bbf56ead82%28Office.15%29.aspx)|
|[CompareEntryIDs](http://msdn.microsoft.com/library/4e935803-9c73-03d2-17c9-dcaf169fdbbe%28Office.15%29.aspx)|
|[CreateContactCard](http://msdn.microsoft.com/library/d050e0e3-3c0d-bd01-f008-2628056625d1%28Office.15%29.aspx)|
|[CreateRecipient](http://msdn.microsoft.com/library/7134c0d7-5f60-c63c-2dde-492d52b78fbe%28Office.15%29.aspx)|
|[CreateSharingItem](http://msdn.microsoft.com/library/4c93d347-cc39-eb5d-bf08-125b69f91eb6%28Office.15%29.aspx)|
|[Dial](http://msdn.microsoft.com/library/1fd29ed8-e983-c668-c48f-f642c56bfcd2%28Office.15%29.aspx)|
|[GetAddressEntryFromID](http://msdn.microsoft.com/library/04e9d2c5-231d-35c8-eafa-0e58fbd7a2a1%28Office.15%29.aspx)|
|[GetDefaultFolder](http://msdn.microsoft.com/library/761b8b53-dd4d-43e4-c8f0-69cefdf0c77a%28Office.15%29.aspx)|
|[GetFolderFromID](http://msdn.microsoft.com/library/0fb2d3b5-2967-1943-922a-7ec03e514e62%28Office.15%29.aspx)|
|[GetGlobalAddressList](http://msdn.microsoft.com/library/0c892483-96c5-461d-a862-fe84ddcce097%28Office.15%29.aspx)|
|[GetItemFromID](http://msdn.microsoft.com/library/f2abff80-4c04-998b-654b-28600424a16f%28Office.15%29.aspx)|
|[GetRecipientFromID](http://msdn.microsoft.com/library/8475e869-ce1f-cd10-0c02-79a6dd5f9a8e%28Office.15%29.aspx)|
|[GetSelectNamesDialog](http://msdn.microsoft.com/library/883d90e0-b3cc-e76e-cbe6-cb271e9ccb37%28Office.15%29.aspx)|
|[GetSharedDefaultFolder](http://msdn.microsoft.com/library/e2196423-e4f2-2797-c16c-dc54e2c0f7d2%28Office.15%29.aspx)|
|[GetStoreFromID](http://msdn.microsoft.com/library/ba5b3df8-22a5-39fa-68ab-9f1e4cfe7f47%28Office.15%29.aspx)|
|[Logoff](http://msdn.microsoft.com/library/f9b15e80-a942-3d76-63ef-00c0a140337d%28Office.15%29.aspx)|
|[Logon](http://msdn.microsoft.com/library/167c632b-0d52-a1e4-8dcd-57d301cde3c9%28Office.15%29.aspx)|
|[OpenSharedFolder](http://msdn.microsoft.com/library/907efeab-8a37-98a6-f241-0a051f03f472%28Office.15%29.aspx)|
|[OpenSharedItem](http://msdn.microsoft.com/library/ebfed85c-0af5-eb72-7a58-ae9e8b655347%28Office.15%29.aspx)|
|[PickFolder](http://msdn.microsoft.com/library/f5c1f35a-8e77-8e7f-fcbe-30c6bc90287a%28Office.15%29.aspx)|
|[RemoveStore](http://msdn.microsoft.com/library/4353387a-0e44-1d4a-b0e6-96e2c2594a6d%28Office.15%29.aspx)|
|[SendAndReceive](http://msdn.microsoft.com/library/196b15b0-6766-ca2a-8dbe-991fc93b8307%28Office.15%29.aspx)|

## Properties



|**Name**|
|:-----|
|[Accounts](http://msdn.microsoft.com/library/80e969ea-d2cc-966d-5fe4-68d59951b5c9%28Office.15%29.aspx)|
|[AddressLists](http://msdn.microsoft.com/library/68b236db-f964-6f7f-6246-e79c6ada19e9%28Office.15%29.aspx)|
|[Application](http://msdn.microsoft.com/library/c7730473-4109-4052-08eb-7cd7d3c1909f%28Office.15%29.aspx)|
|[AutoDiscoverConnectionMode](http://msdn.microsoft.com/library/a73a71ca-0f40-3c7e-bb89-9d6a45775c6f%28Office.15%29.aspx)|
|[AutoDiscoverXml](http://msdn.microsoft.com/library/34834000-1f53-2bfb-7546-886c6e2716fd%28Office.15%29.aspx)|
|[Categories](http://msdn.microsoft.com/library/3963afca-3a7e-38d7-1347-7e1467be3a10%28Office.15%29.aspx)|
|[Class](http://msdn.microsoft.com/library/de558f45-5c09-7285-39cd-8c8525eb28ec%28Office.15%29.aspx)|
|[CurrentProfileName](http://msdn.microsoft.com/library/731df710-cb42-eb68-8fbc-790b74468491%28Office.15%29.aspx)|
|[CurrentUser](http://msdn.microsoft.com/library/d6884fcf-c1de-23f4-8d91-02c8f9fd5253%28Office.15%29.aspx)|
|[DefaultStore](http://msdn.microsoft.com/library/4080e227-bd76-3168-7bc7-93fe04023a3b%28Office.15%29.aspx)|
|[ExchangeConnectionMode](http://msdn.microsoft.com/library/4b9f7917-5340-cf72-d690-ac5a7b8d4792%28Office.15%29.aspx)|
|[ExchangeMailboxServerName](http://msdn.microsoft.com/library/027d8d2d-612d-8eda-a6af-aa8dd371013e%28Office.15%29.aspx)|
|[ExchangeMailboxServerVersion](http://msdn.microsoft.com/library/01e83a30-f574-1ff6-34de-85c14ecc09c1%28Office.15%29.aspx)|
|[Folders](http://msdn.microsoft.com/library/a732d338-c825-4d38-0107-345069da708c%28Office.15%29.aspx)|
|[Offline](http://msdn.microsoft.com/library/c62112d5-e50f-bd6a-bb3b-7c1818752d8b%28Office.15%29.aspx)|
|[Parent](http://msdn.microsoft.com/library/d03ca579-3739-a8ef-fda7-650aa7d7d2d1%28Office.15%29.aspx)|
|[Session](http://msdn.microsoft.com/library/93dba2e5-d11e-7761-ac29-08f5b7a83b49%28Office.15%29.aspx)|
|[Stores](http://msdn.microsoft.com/library/4ffdc2b3-be7b-da21-ac85-bde5481ae2f2%28Office.15%29.aspx)|
|[SyncObjects](http://msdn.microsoft.com/library/0948f154-022f-b12e-87e3-1b3a4ce127c3%28Office.15%29.aspx)|
|[Type](http://msdn.microsoft.com/library/a6872028-0588-94b6-086a-03cf830cd339%28Office.15%29.aspx)|

## See also


#### Other resources


[NameSpace Object Members](http://msdn.microsoft.com/library/d7a978a3-a2c8-6195-c5f8-af8773500456%28Office.15%29.aspx)
[How to: Obtain and Log On to an Instance of Outlook](http://msdn.microsoft.com/library/ef369364-6500-2759-3ef4-ed4411112e96%28Office.15%29.aspx)
[Outlook Object Model Reference](http://msdn.microsoft.com/library/73221b13-d8d8-99b8-3394-b95dbbfd5ddc%28Office.15%29.aspx)
