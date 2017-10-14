---
title: Using Multiple Accounts for the Same Profile on Outlook
ms.prod: outlook
ms.assetid: 9e06e076-d62a-37c8-4502-709da5a0b104
ms.date: 06/08/2017
---


# Using Multiple Accounts for the Same Profile on Outlook

You can sign in to Outlook by using a profile that specifies one or more accounts associated with different delivery stores. For a given session, the **NameSpace** object has members that maintain and access information for the primary Exchange account, and the [Accounts](namespace-accounts-property-outlook.md) property of the [NameSpace](namespace-object-outlook.md) object holds information for all the accounts defined for the session's profile. 

The **NameSpace.Accounts** property returns an [Accounts](accounts-object-outlook.md) collection for the current profile, tracking information for all accounts including Exchange, IMAP, and POP3 accounts, each of which can be associated with a different delivery store. To identify the primary Exchange account in the **Accounts** collection for a session, look for the account that has the [ExchangeStoreType](store-exchangestoretype-property-outlook.md) property of the store (that is specified by **[Account.DeliveryStore](account-deliverystore-property-outlook.md)**) equal to  **OlExchangeStoreType.olPrimaryExchangeMailbox**.

```
Account.DeliveryStore.ExchangeStoreType = OlExchangeStoreType.olPrimaryExchangeMailbox
```

The following table compares members of the **NameSpace** object and members of the [Account](account-object-outlook.md), **Accounts**, or [Store](store-object-outlook.md) object depending on whether the session's profile has one account or multiple accounts. If only the primary Exchange account is in the session's profile, use the following members of the NameSpace object. 

|**Description**|**Purpose**|
|:-----|:-----|
|Use the following members of the noted objects if there are multiple accounts in the session's profile.|**[AutoDiscoverConnectionMode](namespace-autodiscoverconnectionmode-property-outlook.md)** property, **[AutoDiscoverXml](namespace-autodiscoverxml-property-outlook.md)** property, **[AutoDiscoverComplete](namespace-autodiscovercomplete-event-outlook.md)** event|
|To use auto discovery of the Exchange server that hosts the primary Exchange account mailbox.|**[Account.AutoDiscoverConnectionMode](account-autodiscoverconnectionmode-property-outlook.md)** property, **[Account.AutoDiscoverXml](account-autodiscoverxml-property-outlook.md)** property, **[Accounts.AutoDiscoverComplete](accounts-autodiscovercomplete-event-outlook.md)** event|
|To use auto discovery of the Exchange server that hosts the account mailbox.|**[ExchangeConnectionMode](namespace-exchangeconnectionmode-property-outlook.md)** property, **[ExchangeMailboxServerName](namespace-exchangemailboxservername-property-outlook.md)** property, **[ExchangeMailboxServerVersion](namespace-exchangemailboxserverversion-property-outlook.md)** property|
|To obtain information for the Exchange server that hosts the primary Exchange account mailbox.|**[Account.ExchangeConnectionMode](account-exchangeconnectionmode-property-outlook.md)** property, **[Account.ExchangeMailboxServerName](account-exchangemailboxservername-property-outlook.md)** property, **[Account.ExchangeMailboxServerVersion](account-exchangemailboxserverversion-property-outlook.md)** property
|To obtain information for the Exchange server that hosts the account mailbox.|**[Categories](namespace-categories-property-outlook.md)** property|
|To obtain a **[Categories](categories-object-outlook.md)** collection that represents the Master Category List for the primary account of the session.|**[Store.Categories](store-categories-property-outlook.md)** property|
|To obtain a [Categories](categories-object-outlook.md) collection that represents the categories defined for the store that is associated with an account in the session's profile.|**[CurrentUser](namespace-currentuser-property-outlook.md)** property|
|To obtain a **[Recipient](recipient-object-outlook.md)** object that represents the user currently logged on for the session.|**[Account.CurrentUser](account-currentuser-property-outlook.md)** property|
|To obtain a **Recipient** object that represents the user of the account that is defined in the session's profile. The account can be any account that Outlook supports including Exchange, IMAP, and POP3.|**[DefaultStore](namespace-defaultstore-property-outlook.md)** property|
|To obtain the default store for the session's profile.| **[Account.DeliveryStore](account-deliverystore-property-outlook.md)** property|
|To obtain the default delivery store for the account that is defined in the session's profile. The account can be any account that Outlook supports including Exchange, IMAP, and POP3.|**[GetAddressEntryFromID](namespace-getaddressentryfromid-method-outlook.md)** method|
|To obtain an **[AddressEntry](addressentry-object-outlook.md)** object that corresponds to the given entry ID.|**[Account.GetAddressEntryFromID](account-getaddressentryfromid-method-outlook.md)** method|
|To obtain an **AddressEntry** object that corresponds to the account and given entry ID. The account can be any account that Outlook supports including Exchange, IMAP, and POP3.|**[GetRecipientFromID](namespace-getrecipientfromid-method-outlook.md)** method|
|To obtain a **Recipient** object that corresponds to the given entry ID.|**[Account.GetRecipientFromID](account-getrecipientfromid-method-outlook.md)** method|
|To obtain a **Recipient** object that corresponds to the account and given entry ID. |The account can be any account that Outlook supports including Exchange, IMAP, and POP3.|

If you are operating with multiple accounts in the current profile, see the following tasks:

-  [How to: Obtain Information for Multiple Accounts](obtain-information-for-multiple-accounts.md)
    
-  [How to: Identify a Folder with an Account](identify-a-folder-with-an-account.md)
    
-  [How to: Create a Sendable Item for a Specific Account Based on the Current Folder](create-a-sendable-item-for-a-specific-account-based-on-the-current-folder-outloo.md)
    
-  [How to: Identify a Global Address List or a Set of Address Lists with a Store](identify-the-global-address-list-or-a-set-of-address-lists-with-a-store.md)
    

