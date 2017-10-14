---
title: Identify the Global Address List or a Set of Address Lists with a Store
ms.prod: outlook
ms.assetid: 2cca6dc2-883d-b8cf-cd60-98614d2fb673
ms.date: 06/08/2017
---


# Identify the Global Address List or a Set of Address Lists with a Store

In a Microsoft Outlook session where multiple Microsoft Exchange accounts are defined in the profile, there can be multiple address lists associated with a store. This topic has two code samples that show how to retrieve the Global Address List for a given store, and how to obtain all of the  [AddressList](addresslist-object-outlook.md) objects associated with a given store. In each of these code samples, the specific store of interest is the store for the current folder displayed in the active explorer, but the algorithm to get a Global Address List or a set of address lists for a store applies to any Exchange store.

The following managed code is written in C#. To run a .NET Framework managed code sample that needs to call into a Component Object Model (COM), you must use an interop assembly that defines and maps managed interfaces to the COM objects in the object model type library. For Outlook, you can use Visual Studio and the Outlook Primary Interop Assembly (PIA). Before you run managed code samples for Outlook 2013, ensure that you have installed the Outlook 2013 PIA and have added a reference to the Microsoft Outlook 15.0 Object Library component in Visual Studio. You should use the following code in the  `ThisAddIn` class of an Outlook add-in (using Office Developer Tools for Visual Studio). The **Application** object in the code must be a trusted Outlook **Application** object provided by `ThisAddIn.Globals`. For more information about using the Outlook PIA to develop managed Outlook solutions, see the  **Welcome to the Outlook Primary Interop Assembly Reference** on MSDN.

The first code sample contains the  `DisplayGlobalAddressListForStore` method and the `GetGlobalAddressList` function. The `DisplayGlobalAddressListForStore` method displays the Global Address List that is associated with the current store in the **Select Names** dialog box. `DisplayGlobalAddressListForStore` first obtains the current store. If the current store is an Exchange store, calls `GetGlobalAddressList` to obtain the Global Address List associated with the current store. `GetGlobalAddressList` uses the [PropertyAccessor](propertyaccessor-object-outlook.md) object and the MAPI property, http://schemas.microsoft.com/mapi/proptag/0x3D150102, to obtain the UIDs of an address list and the current store. `GetGlobalAddressList` identifies an address list as associated with a store if their UIDs match, and the address list is the Global Address List if its [AddressListType](addresslist-addresslisttype-property-outlook.md) property is **olExchangeGlobalAddressList**. If the call to  `GetGlobalAddressList` succeeds, `DisplayGlobalAddressListForStore` uses the [SelectNamesDialog](selectnamesdialog-object-outlook.md) object to display the returned Global Address List in the **Select Names** dialog box.




```C#
void DisplayGlobalAddressListForStore() 
{ 
    // Obtain the store for the current folder 
    // as the current store. 
    Outlook.Folder currentFolder = 
        Application.ActiveExplorer().CurrentFolder 
        as Outlook.Folder; 
    Outlook.Store currentStore = currentFolder.Store; 
 
    // Check if the current store is Exchange. 
    if (currentStore.ExchangeStoreType != 
        Outlook.OlExchangeStoreType.olNotExchange) 
    { 
        Outlook.SelectNamesDialog snd =  
            Application.Session.GetSelectNamesDialog(); 
 
        // Try to get the Global Address List associated  
        // with the current store. 
        Outlook.AddressList addrList =  
            GetGlobalAddressList(currentStore); 
        if (addrList != null) 
        { 
            // Display the Global Address List in the  
            // Select Names dialog box. 
            snd.InitialAddressList = addrList; 
            snd.Display(); 
        } 
    } 
} 
 
public Outlook.AddressList GetGlobalAddressList(Outlook.Store store) 
{ 
    // Property string for the UID of a store or address list. 
    string  PR_EMSMDB_SECTION_UID =  
        @"http://schemas.microsoft.com/mapi/proptag/0x3D150102"; 
 
    if (store == null) 
    { 
        throw new ArgumentNullException(); 
    } 
 
    // Obtain the store UID using the proprety string and  
    // property accessor on the store. 
    Outlook.PropertyAccessor oPAStore = store.PropertyAccessor; 
 
    // Convert the store UID to a string value. 
    string storeUID = oPAStore.BinaryToString( 
        oPAStore.GetProperty(PR_EMSMDB_SECTION_UID)); 
 
    // Enumerate each address list associated 
    // with the session. 
    foreach (Outlook.AddressList addrList  
        in Application.Session.AddressLists) 
    { 
        // Obtain the address list UID and convert it to  
        // a string value. 
        Outlook.PropertyAccessor oPAAddrList =  
            addrList.PropertyAccessor; 
        string addrListUID = oPAAddrList.BinaryToString( 
            oPAAddrList.GetProperty(PR_EMSMDB_SECTION_UID)); 
 
        // Return the address list associated with the store 
        // if the address list UID matches the store UID and 
        // type is olExchangeGlobalAddressList. 
        if (addrListUID == storeUID &;&; addrList.AddressListType == 
            Outlook.OlAddressListType.olExchangeGlobalAddressList) 
        { 
            return addrList; 
        } 
    } 
    return null; 
} 

```

The second code sample contains the  `EnumerateAddressListsForStore` method and `GetAddressLists` function. The `EnumerateAddressListsForStore` method displays the type and resolution order of each address list defined for the current store. `EnumerateAddressListsForStore` first obtains the current store, then it calls `GetAddressLists` to obtain a .NET Framework generic **List** object that contains **AddressList** objects for the current store. `GetAddressLists` enumerates each address list defined for the session, uses the [PropertyAccessor](propertyaccessor-object-outlook.md) object and the MAPI named property, http://schemas.microsoft.com/mapi/proptag/0x3D150102, to obtain the UIDs of an address list and the current store. `GetGlobalAddressList` identifies an address list as associated with a store if their UIDs match, and returns a set of address lists for the current store. `EnumerateAddressListsForStore` then uses the [AddressListType](addresslist-addresslisttype-property-outlook.md) and [ResolutionOrder](addresslist-resolutionorder-property-outlook.md) properties of the **AddressList** object to display the type and resolution order for each returned address list.



```C#
private void EnumerateAddressListsForStore() 
{ 
    // Obtain the store for the current folder 
    // as the current store. 
    Outlook.Folder currentFolder = 
       Application.ActiveExplorer().CurrentFolder 
       as Outlook.Folder; 
    Outlook.Store currentStore = currentFolder.Store; 
 
    // Obtain all address lists for the current store. 
    List<Outlook.AddressList> addrListsForStore =  
        GetAddressLists(currentStore); 
    foreach (Outlook.AddressList addrList in addrListsForStore) 
    { 
        // Display the type and resolution order of each  
        // address list in the current store. 
        Debug.WriteLine(addrList.Name  
            + " " + addrList.AddressListType.ToString() 
            + " Resolution Order: " + 
            addrList.ResolutionOrder); 
     }  
} 
 
public List<Outlook.AddressList> GetAddressLists(Outlook.Store store) 
{ 
    List<Outlook.AddressList> addrLists =  
        new List<Microsoft.Office.Interop.Outlook.AddressList>(); 
 
    // Property string for the UID of a store or address list. 
    string PR_EMSMDB_SECTION_UID = 
        @"http://schemas.microsoft.com/mapi/proptag/0x3D150102"; 
 
    if (store == null) 
    { 
        throw new ArgumentNullException(); 
    } 
 
    // Obtain the store UID and convert it to a string value. 
    Outlook.PropertyAccessor oPAStore = store.PropertyAccessor; 
    string storeUID = oPAStore.BinaryToString( 
        oPAStore.GetProperty(PR_EMSMDB_SECTION_UID)); 
 
    // Enumerate each address list associated 
    // with the session. 
    foreach (Outlook.AddressList addrList 
        in Application.Session.AddressLists) 
    { 
        // Obtain the address list UID and convert it to  
        // a string value. 
        Outlook.PropertyAccessor oPAAddrList = 
            addrList.PropertyAccessor; 
        string addrListUID = oPAAddrList.BinaryToString( 
            oPAAddrList.GetProperty(PR_EMSMDB_SECTION_UID)); 
         
        // Add the address list to the resultant set of address lists 
        // if the address list UID matches the store UID. 
        if (addrListUID == storeUID) 
        { 
            addrLists.Add(addrList); 
        } 
    } 
    // Return the set of address lists associated with the store. 
    return addrLists; 
} 

```


