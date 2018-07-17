---
title: Protected Properties and Methods
ms.prod: outlook
ms.assetid: 8522d350-a257-2924-2260-3cc02b6ebbca
ms.date: 06/08/2017
---


# Protected Properties and Methods

This topic lists the properties and methods in the Outlook object model that are protected by the Object Model Guard. If untrusted code performs a get on these properties or uses these methods, under default conditions for how Outlook is set up, it will invoke a security warning. The user will then have to verify and respond to the warning in order to proceed.

There are three security prompts that an untrusted application can possibly invoke, depending on the protected property or method that the application uses:

- The address book warning. This is the most common of the three security prompts. Unless marked otherwise, the properties and methods in the table below generate the address book warning.
    
- The execute action warning. Properties and methods superscripted by 1 in the table below denote that they generate the execute action warning.
    
- The send message warning. Properties and methods superscripted by 2 in the table below denote that they generate the send message warning.
    
For more information on security warnings, see  [Outlook Object Model Security Prompts](outlook-object-model-security-warnings.md).



| **Object**| **Protected Properties**| **Protected Methods**|
|:-----|:-----|:-----|
| [Account](account-object-outlook.md)|CurrentUser, SmtpAddress|GetAddressEntryFromID, GetRecipientFromID|
| [Action](action-object-outlook.md)||Execute1|
| [AddressEntries](addressentries-object-outlook.md)||Add, GetFirst, GetLast, GetNext, GetPrevious, Item|
| [AddressEntry](addressentry-object-outlook.md)|Address, ID, Manager, Members, Parent, PropertyAccessor|GetExchangeDistributionList, GetExchangeUser, Update|
| [AddressList](addresslist-object-outlook.md)|AddressEntries, ID, PropertyAccessor||
| [AddressLists](addresslists-object-outlook.md)||Item|
| [AppointmentItem](appointmentitem-object-outlook.md)|Body, OptionalAttendees, Organizer, PropertyAccessor, RequiredAttendees, Resources, RTFBody|Respond2, SaveAs, Send2|
| [Attachment](attachment-object-outlook.md)|PropertyAccessor||
| [CalendarSharing](calendarsharing-object-outlook.md)||SaveAsICal|
| [Columns](columns-object-outlook.md)||Add|
| [ContactItem](contactitem-object-outlook.md)|Body, Email1Address, Email1AddressType, Email1DisplayName, Email1EntryID, Email2Address, Email2AddressType, Email2DisplayName, Email2EntryID, Email3Address, Email3AddressType, Email3DisplayName, Email3EntryID, IMAddress, NetMeetingAlias, PropertyAccessor, ReferredBy, RTFBody|SaveAs|
| [DistListItem](distlistitem-object-outlook.md)|Body, PropertyAccessor, RTFBody|GetMember, SaveAs|
| [DocumentItem](documentitem-object-outlook.md)|Body, PropertyAccessor||
| [ExchangeDistributionList](exchangedistributionlist-object-outlook.md)|Address, Alias, ID, Parent, PrimarySmtpAddress, PropertyAccessor|GetExchangeDistributionList, GetExchangeUser, GetMemberOfList, GetExchangeDistributionListMembers, GetOwners, Update|
| [ExchangeUser](exchangeuser-object-outlook.md)|Address, Alias, ID, Parent, PrimarySmtpAddress, PropertyAccessor|GetDirectReports, GetExchangeDistributionList, GetExchangeUser, GetExchangeUserManager, GetMemberOfList, Update|
| [Folder](folder-object-outlook.md)|GetCalendarExporter, PropertyAccessor||
| [Inspector](inspector-object-outlook.md)|HTMLEditor, WordEditor||
| [ItemProperties](itemproperties-object-outlook.md)|Any protected property for an item||
| [JournalItem](journalitem-object-outlook.md)|Body, ContactNames, PropertyAccessor|SaveAs|
| [MailItem](mailitem-object-outlook.md)|Bcc, Body, Cc, HTMLBody, PropertyAccessor, ReceivedByName, ReceivedOnBehalfOfName, Recipients, ReplyRecipientNames, RTFBody, Sender, SenderEmailAddress, SenderEmailType, SenderName, SentOnBehalfOfName, To|SaveAs, Send2|
| [MeetingItem](meetingitem-object-outlook.md)|Body, PropertyAccessor, Recipients, RTFBody, SenderName|SaveAs|
| [MobileItem](http://msdn.microsoft.com/library/da8149d5-66d3-ea02-941f-e7f2f9eb6bc3%28Office.15%29.aspx)|Body, HTMLBody, PropertyAccessor, ReceivedByName, Recipients, ReplyRecipientNames, SenderEmailAddress, SenderEmailType, SenderName, SMILBody, To|SaveAs, Send2|
| [NameSpace](namespace-object-outlook.md)|CurrentUser, SelectNamesDialog|GetAddressEntryFromID, GetRecipientFromID|
| [NoteItem](noteitem-object-outlook.md)|Body, PropertyAccessor||
| [PostItem](postitem-object-outlook.md)|Body, HTMLBody, PropertyAccessor, RTFBody, SenderName|SaveAs|
| [Recipient](recipient-object-outlook.md)|Any property|Any method|
| [Recipients](recipients-object-outlook.md)|Any property|Any method|
| [RemoteItem](remoteitem-object-outlook.md)|Body, PropertyAccessor||
| [ReportItem](reportitem-object-outlook.md)|Body, PropertyAccessor||
| [SelectNamesDialog](selectnamesdialog-object-outlook.md)|Recipients||
| [SharingItem](sharingitem-object-outlook.md)|Bcc, Body, Cc, HTMLBody, PropertyAccessor, ReceivedByName, ReceivedOnBehalfOfName, ReplyRecipientNames, RTFBody, SenderEmailAddress, SenderEmailType, SenderName, SendOnBehalfOfName, To|Allow, SaveAs, Send2|
| [StorageItem](storageitem-object-outlook.md)|Body, PropertyAccessor||
| [Store](store-object-outlook.md)|PropertyAccessor||
| [TaskItem](taskitem-object-outlook.md)|Body, ContactNames, Contacts, Delegator, Owner, PropertyAccessor, RTFBody, StatusOnCompletionRecipients, StatusUpdateRecipients|SaveAs, Send2|
| [TaskRequestAcceptItem](taskrequestacceptitem-object-outlook.md)|Body, PropertyAccessor, RTFBody||
| [TaskRequestDeclineItem](taskrequestdeclineitem-object-outlook.md)|Body, PropertyAccessor, RTFBody||
| [TaskRequestItem](taskrequestitem-object-outlook.md)|Body, PropertyAccessor, RTFBody||
| [TaskRequestUpdateItem](taskrequestupdateitem-object-outlook.md)|Body, PropertyAccessor, RTFBody||
| [UserProperties](userproperties-object-outlook.md)||Find|
| [UserProperty](userproperty-object-outlook.md)|Formula||


 **Note**   **[UserProperties.Find](userproperties-find-method-outlook.md)** is protected if the property being requested is one of the built-in properties that contains address information. If you ask for a custom property or a property like **Subject** that doesn't contain address information, a prompt will not be displayed.


