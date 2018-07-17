---
title: Sharing Online Calendars, RSS Feeds, Microsoft SharePoint Foundation Folders, and Exchange Folders
ms.prod: outlook
ms.assetid: e579e026-bd10-37bb-eb3e-5c9f042fa0fa
ms.date: 06/08/2017
---


# Sharing Online Calendars, RSS Feeds, Microsoft SharePoint Foundation Folders, and Exchange Folders

 In Microsoft Outlook, you can share and subscribe to a variety of online resources, including:


- Webcal calendars (webcal:// _mysite_/ _mycalendar_)
    
- RSS feeds (feed:// _mysite_/ _myfeed_)
    
- SharePoint Foundation folders (stssync:// _mysite_/ _myfolder_)
    
- Exchange folders
    

Calendar information can also be shared either by providing direct access to a calendar folder or by exporting calendar information to an iCalendar calendar (.ics) file. For more information about sharing calendars, see  [Sharing Calendars](sharing-calendars.md).


## Sharing Online Resources

For publically available online resources, such as Webcal calendars, RSS feeds, and SharePoint Foundation folders, a sharing message is not required. You can use the  **[OpenSharedFolder](namespace-opensharedfolder-method-outlook.md)** method of the **NameSpace** object to open the online resource. For online resources to which access is required, such as Exchange folders, a sharing request can be created to request access. You can create a sharing request by using the **[CreateSharingItem](namespace-createsharingitem-method-outlook.md)** method of the **[NameSpace](namespace-object-outlook.md)** object to create a **[SharingItem](sharingitem-object-outlook.md)** object. The shared resource (a **[Folder](folder-object-outlook.md)** object reference to the desired Exchange default folder) is used to establish the sharing context for the sharing request.

You can also use a sharing invitation to direct another user to an online resource. To construct a sharing invitation, the  **CreateSharingItem** method of the **[NameSpace](namespace-object-outlook.md)** object is used to create a **SharingItem** object. The shared resource (either a **[Folder](folder-object-outlook.md)** object reference to the desired folder or a string containing the appropriate URI for the online resource) is used to establish the sharing context for the sharing invitation.


 **Note**  Sharing requests can be created only for Exchange default folders. To access other Exchange folders, a sharing invitation from the owner of the folder is required.


## Sharing Providers

Each type of online resource, such as Webcal calendars, is supported by a corresponding sharing provider. A sharing provider encapsulates the access and interpretation tools for a given online resource type. You can use the  **[SharingProvider](sharingitem-sharingprovider-property-outlook.md)** and **[SharingProviderGuid ](sharingitem-sharingproviderguid-property-outlook.md)** properties of a **SharingItem** to determine the sharing provider used by a given sharing message.


