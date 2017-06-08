---
title: Storing Outlook Items
ms.prod: outlook
ms.assetid: e4a639a4-10b2-7665-9261-19d6e7707e48
ms.date: 06/08/2017
---


# Storing Outlook Items

This topic describes how Outlook items are stored in folders and stores based on an account in the current profile.

The Outlook object model provides the following objects to store Outlook items:

- The  **[Folder](folder-object-outlook.md)** object, which represents a container for other **Folder** objects and Outlook items.
    
     **Note**  The  **Folder** object has replaced the **MAPIFolder** object that existed in Microsoft Office Outlook 2003 and earlier versions of Outlook. New solutions should only use **Folder**.
- The  **[Folders](folders-object-outlook.md)** collection, which represents all the **Folder** objects at one level of the folder tree in a store. The **Folders** collection can also represent a collection of search folders.
    
     **Note**  Although a search folder is represented programmatically by a  **Folder** object, not all events, methods, and properties of **Folder** apply to search folders.
- The  **[Store](store-object-outlook.md)** object, which represents a file on the local computer or a network drive that stores e-mail messages and other items. If you use an Exchange server, you can have a store on the server, in an Exchange Public folder, or on a local computer in a Personal Folders File (.pst) or Offline Folder File (.ost). For a POP3, IMAP, and HTTP e-mail server, a store is a .pst file.
    
    You can add a store to the current profile using  **[NameSpace.AddStore](namespace-addstore-method-outlook.md)** and **[NameSpace.AddStoreEx](namespace-addstoreex-method-outlook.md)**, and remove an existing store from the current profile using  **[NameSpace.RemoveStore](namespace-removestore-method-outlook.md)**.
    
- The  **[Stores](stores-object-outlook.md)** collection, which represents all the stores in the current Outlook profile. A profile defines one or more e-mail accounts, and each e-mail account is associated with a server of a specific type. The type of server determines the type of the store and how e-mail and other items are delivered and stored. For example, an Exchange server stores e-mail and other items in either a .pst file or a .ost file on the local computer or a mapped network drive, and an HTTP server (such as Hotmail) stores items in a .pst file on the local computer.
    

The  **Store** and **Stores** objects support the following:

- Enumerating folders in a store using  **[Store.GetRootFolder](store-getrootfolder-method-outlook.md)** and then **[Folder.Folders](folder-folders-property-outlook.md)**.
    
- Enumerating search folders in a store using  **[Store.GetSearchFolders](store-getsearchfolders-method-outlook.md)**.
    
     **Note**  Since a store does not necessarily support search folders, in general, you should trap for returned errors when using  **Store.GetSearchFolders** to obtain any search folders on a store.
- Better performance with enumerating folders. Because getting the root folder or search folders in a store requires the store to be open and opening a store imposes an overhead on performance, you can check the  **[Store.IsOpen](store-isopen-property-outlook.md)** property before you decide to pursue the operation.
    
- Locating a local store (.pst or .ost) for an Exchange server, or a store (.pst) for a POP3, IMAP, or HTTP e-mail server, using the  **[Store.FilePath](store-filepath-property-outlook.md)** property.
    
- Discovery of the Exchange store type and differentiation among different Exchange store types using the  **[Store.ExchangeStoreType](store-exchangestoretype-property-outlook.md)** property.
    
- Additional information for an Exchange server through the  **[Store.IsCachedExchange](store-iscachedexchange-property-outlook.md)** and **[Store.IsDataFileStore](store-isdatafilestore-property-outlook.md)** properties.
    
- The  **[PropertyAccessor](propertyaccessor-object-outlook.md)** object through the **[Store.PropertyAccessor](store-propertyaccessor-property-outlook.md)** property, allowing access to store properties that are not exposed as explicit built-in properties in the Outlook object model.
    


