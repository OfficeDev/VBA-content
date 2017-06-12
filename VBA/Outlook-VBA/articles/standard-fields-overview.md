---
title: Standard Fields Overview
ms.prod: outlook
ms.assetid: f0d903a3-f404-8511-af3d-d4f3e30f0779
ms.date: 06/08/2017
---


# Standard Fields Overview

Microsoft Outlook provides standard fields for each standard item. Some fields are available in individual items. You can add and remove all fields from a table or card view type.

The value for fields with the  **Yes/No** data type is saved as **-1** or **0**. The value for fields with the  **Duration** data type is saved as minutes.

## What do you want more information about?




### Standard fields in an appointment or meeting request





|**Field**|**Edit in view**|**Data type and meaning**|
|:-----|:-----|:-----|
|All Day Event|Yes| **Yes/No**. If set to  **Yes**, the  **Duration** field is set to 24 hours (1440 minutes).|
|Attachment|No| **Yes/No**.|
|Billing Information|Yes| **Text**.|
|Categories|Yes| **Text**. Field used to group and find related items. Multiple categories are separated by commas.|
|Contacts|No| **Text**. Names of contacts linked to this item. Multiple names are separated by commas.|
|Conversation|No| **Text**. Value of the  **Subject** field in the original message in a conversation.|
|Created|No| **Date/Time**. Date and time the  **Calendar** item is created.|
|Do Not AutoArchive|Yes| **Yes/No**. Specifies whether to archive the  **Calendar** item.|
|Duration|No| **Duration**. 24 hours (1440 minutes) if the  **All Day Event** field is set to **Yes.** Otherwise, the difference between the values of the **End** and **Start** fields. Saved as minutes.|
|End|Yes| **Date/Time**. End date and time of a  **Calendar** item.|
|Importance|Yes|The following settings apply: **0** Low importance **1** Normal importance **2** High importance|
|Icon|Yes|Internal data type.|
|In Folder|No| **Text**. Name of the folder that contains the  **Calendar** item.|
|Location|Yes| **Text**. Location of a meeting or appointment.|
|Meeting Status|No|The following settings apply: **0** None **1** Meeting organizer **2** Tentatively accepted **3** Accepted **4** Declined **5** Not yet accepted|
|Message Class|No|Specifies the message class for the type of item. |
|Mileage|Yes| **Text**.|
|Modified|No| **Date/Time**. Last time the  **Calendar** item was modified.|
|NetMeeting AutoStart|Yes| **Yes/No**. Specifies whether an online meeting starts immediately when the reminder appears.|
|NetMeeting Office Document Path|Yes| **Text**. Specifies the path of the Microsoft Office document used for online-meeting collaboration.|
|NetMeeting Organizer E-mail|Yes| **Text**. E-mail address of the online-meeting organizer.|
|NetMeeting Server|Yes| **Text**. Name of the NetMeeting server for the online meeting.|
|NetMeeting Type|No|The following settings apply: **0** NetMeeting **1** NetShow|
|NetMeeting URL|Yes| **Text**. The URL of the online meeting.|
|Notes|No| **Text**. Value of the text box of the appointment.|
|Online Meeting|Yes|Yes/No.|
|Online Meeting Type|Yes|The following settings apply: **0** NetMeeting **1** NetShow|
|Optional Attendees|No| **Text**. Names of optional attendees for a meeting or appointment. Multiple names are separated by semicolons.|
|Organizer|No| **Text**. Name of the organizer of a meeting or appointment.|
|Outlook Internal Version|No|For administrator use only.|
|Outlook Version|No| **Text**. Version of Outlook that the  **Calendar** item is created in.|
|Read|No| **Yes/No**. Specifies whether the  **Calendar** item has been marked as read.|
|Recurrence|No|The following settings apply: **0** None **1** Daily **2** Weekly **3** Monthly **4** Yearly|
|Recurrence Pattern|No| **Text**. Combination of the values of the  **Recurrence**,  **Start**, and  **End** fields.|
|Recurrence Range End|No| **Date/Time**. Last date and time of a recurring  **Calendar** item.|
|Recurrence Range Start|No| **Date/Time**. First date and time of a recurring  **Calendar** item.|
|Recurring|No| **Yes/No**. Specifies whether the  **Calendar** item recurs.|
|Remind Beforehand|Yes| **Number**. Minutes before the reminder runs prior to a meeting or appointment.|
|Reminder|Yes| **Yes/No**. If the start time for the meeting or appointment has already passed, the  **Reminder** field cannot be set.|
|Reminder Override Default|Yes| **Yes/No**. If set to  **Yes**, the  **Remind Beforehand**,  **Reminder Sound**, and  **Reminder Sound File** fields are used to control the reminder for the item. If set to **No**, the reminders options on the  **Advanced** tab of the **Outlook Options** dialog are used. (To open the **Outlook Options** dialog box, click the **File** tab and then click **Options**.)|
|Reminder Sound|Yes| **Yes/No**. Specifies whether to play the sound file as a reminder.|
|Reminder Sound File|Yes| **Text**. Path of the sound file to be played as a reminder.|
|Required Attendees|No| **Text**. Names of required attendees for a meeting or appointment. Multiple names are separated by semicolons.|
|Resources|No| **Text**. Names of resources for a meeting or appointment. Multiple names are separated by semicolons.|
|Response Requested|Yes| **Yes/No**. In a meeting request, specifies whether the recipient has been asked to respond.|
|Sensitivity|No|The following settings apply: **0** Normal **1** Personal **2** Private **3** Confidential|
|Show Time As|Yes|The following settings apply: **0** Free **1** Tentative **2** Busy **3** Out of Office|
|Size|No| **Number**. Number of bytes used by the  **Calendar** item.|
|Start|Yes| **Date/Time**. Start time of a  **Calendar** item.|
|Subject|Yes| **Text**.|

### Standard fields in a contact





|**Field**|**Edit in view**|**Data type and meaning**|
|:-----|:-----|:-----|
|Account|Yes| **Text**.|
|Address Selected|Yes|Displays the address text that was entered into the  **Address** field based on the value of the **Address Selector** field.|
|Address Selector|Yes|The following settings apply:Home BusinessOther|
|Anniversary|Yes| **Date/Time**. When the  **Anniversary** field has a value, a **Calendar** item is attached to the contact, and the **Attachment** field is set to Yes.|
|Assistant's Name|Yes| **Text**.|
|Assistant's Phone|Yes| **Text**.|
|Attachment|No| **Yes/No**. Set to Yes when the  **Anniversary** or **Birthday** field is a non-empty field.|
|Billing Information|Yes| **Text**.|
|Birthday|Yes| **Date/Time**. When the  **Birthday** field has a value, a **Calendar** item is attached to the contact, and the **Attachment** field is set to Yes.|
|Business Address|Yes (Card view only)| **Text**.|
|Business Address City|Yes| **Text**.|
|Business Address Country|Yes| **Text**.|
|Business Address PO Box|Yes| **Text**.|
|Business Address Postal Code|Yes| **Text**.|
|Business Address State|Yes| **Text**.|
|Business Address Street|Yes (Card view only)| **Text**.|
|Business Fax|Yes| **Text**.|
|Business Home Page|Yes| **Text**.|
|Business Phone|Yes| **Text**.|
|Business Phone 2|Yes| **Text**.|
|Callback|Yes| **Text**.|
|Car Phone|Yes| **Text**.|
|Categories|Yes| **Text**. Field used to group and find related items. Multiple categories are separated by commas.|
|Children|Yes| **Text**.|
|City|Yes| **Text**.|
|Company|Yes| **Text**.|
|Company Main Phone|Yes| **Text**.|
|Computer Network Name|Yes| **Text**.|
|Contacts|No| **Text**. Names of contacts linked to this item. Multiple names are separated by commas.|
|Country/Region|Yes| **Text**.|
|Created|No| **Date/Time**. The date and time the contact is created.|
|Customer ID|Yes| **Text**.|
|Department|Yes| **Text**.|
|E-mail|No| **Text**.|
|E-mail 2|No| **Text**.|
|E-mail 3|No| **Text**.|
|E-mail Selected|Yes|Displays the e-mail address that was entered based on the value of the  **E-mail Selector** field.|
|E-mail Selector|Yes|The following settings apply: **0** E-mail **1** E-mail 2 **2** E-mail 3|
|E-mail Display As|Yes| **Text**. Alternate text that represents the e-mail address stored in the  **E-mail** field. This text displays on the **To** line when addressing a message or appointment.|
|E-mail2 Display As|Yes| **Text**. Alternate text that represents the e-mail address stored in the  **E-mail 2** field. This text displays on the **To** line when addressing a message or appointment.|
|E-mail3 Display As|Yes| **Text**. Alternate text that represents the e-mail address stored in the  **E-mail 3** field. This text displays on the **To** line when addressing a message or appointment.|
|File As|No| **Text**. Value of the  **Full Name** field, unless modified by the user.|
|First Name|Yes| **Text**.|
|Follow Up Flag|Yes| **Text**.|
|FTP Site|Yes| **Text**. FTP site name for the contact.|
|Full Name|Yes| **Text**. Value of the  **Title**,  **First**,  **Middle**,  **Last**, and  **Suffix** fields of the item, separated by spaces. Any changes made to the **Full Name** field are reflected in its component fields.|
|Gender|Yes|The following settings apply: **0** Unspecified **1** Female **2** Male|
|Government ID Number|Yes| **Text**.|
|Hobbies|Yes| **Text**.|
|Home Address|Yes (Card view only)| **Text**.|
|Home Address City|Yes| **Text**.|
|Home Address Country|Yes| **Text**.|
|Home Address PO Box|Yes| **Text**.|
|Home Address Postal Code|Yes| **Text**.|
|Home Address State|Yes| **Text**.|
|Home Address Street|Yes (Card view only)| **Text**.|
|Home Fax|Yes| **Text**.|
|Home Phone|Yes| **Text**.|
|Home Phone 2|Yes| **Text**.|
|Icon|Yes|Internal Data Type.|
|IM Address|Yes| **Instant Messaging address**. (Instant Messaging is a feature of the Microsoft MSN Messenger Service and Microsoft Exchange Instant Messaging Service.)|
|In Folder|Yes| **Text**. Name of the folder that contains the contact.|
|Initials|Yes| **Text**.|
|Internet Free-Busy Address|Yes| **Text**. Refers to the reading and publishing of a calendar user's free/busy map of events. The map is retrieved when a user plans a meeting.|
|ISDN|Yes| **Text**. Phone number for ISDN connection.|
|Job Title|Yes| **Text**.|
|Journal|Yes| **Yes/No**. Specifies whether activities are automatically recorded in the  **Journal** for the contact.|
|Language|Yes| **Text**.|
|Last Name|Yes| **Text**.|
|Location|Yes| **Text**.|
|Mailing Address|Yes (Card view only)| **Text**.|
|Mailing Address Indicator|Yes| **Yes/No**. Indicates whether the address specified by the  **Address Selector** field is the mailing address.|
|Manager's Name|Yes| **Text**.|
|Message Class|| **Text**. Specifies the message class for the type of item. |
|Middle Name|Yes| **Text**.|
|Mileage|Yes| **Text**.|
|Mobile Phone|Yes| **Text**.|
|Modified|No| **Date/Time**. Last date and time the contact was modified.|
|Nickname|Yes| **Text**.|
|Notes|No| **Text**. Value of the text box of the contact.|
|Office Location|Yes| **Text**.|
|Organizational ID Number|Yes| **Text**.|
|Other Address|Yes (Card view only)| **Text**.|
|Other Address City|Yes| **Text**.|
|Other Address Country|Yes| **Text**.|
|Other Address PO Box|Yes| **Text**.|
|Other Address Postal Code|Yes| **Text**.|
|Other Address State|Yes| **Text**.|
|Other Address Street|Yes (Card view only)| **Text**.|
|Other Fax|Yes| **Text**.|
|Other Phone|Yes| **Text**.|
|Outlook Internal Version|No|For administrator use only|
|Outlook Version|No| **Text**. Version of Outlook that the contact is created in.|
|Pager|Yes| **Text**.|
|Personal Home Page|Yes| **Text**.|
|Phone 1 Selected (through Phone 8 Selected)||Displays the phone number that was selected in the corresponding  **Phone Selector** field.|
|Phone 1 Selector (through Phone 8 Selector)||The following settings apply: Business, Home, Business Fax, Mobile, Radio, Car, Other, and ISDN|
|PO Box|Yes| **Text**.|
|Primary Phone|Yes| **Text**.|
|Private|Yes| **Yes/No**. Indicates whether a specific contact is visible to others who have access to the  **Contacts** folder.|
|Profession|Yes| **Text**.|
|Radio Phone|Yes| **Text**.|
|Read|No| **Yes/No**. Specifies whether the contact has been marked as read.|
|Referred By|Yes| **Text**.|
|Reminder|Yes| **Yes/No**.|
|Reminder Time|Yes| **Date/Time**. Date and time that a reminder is run for a contact.|
|Reminder Topic|Yes| **Text**. Caption displayed with reminder flag.|
|Sensitivity|No|The following settings apply: **0** Normal **1** Personal **2** Private **3** Confidential|
|Size|No| **Number**. Number of bytes used by the contact.|
|Send Plain Text Only|Yes| **Text**.|
|Spouse|Yes| **Text**.|
|State|Yes| **Text**.|
|Street Address|Yes (Card view only)| **Text**.|
|Subject|Yes| **Text**. Value of the  **Full Name** field. If the **Full Name** field is empty, the value of the **File As** field is used.|
|Suffix|Yes| **Text**. Name suffix, such as Jr. or Ph.D.|
|Telex|Yes| **Text**.|
|Title|Yes| **Text**. Name title, such as Mr., Ms., or Mrs.|
|TTY/TDD Phone|Yes| **Text**. Phone number for TTY/TDD connection.|
|User Field 1 - 4|Yes| **Text**. User-defined fields provided for compatibility with other programs. |
|Web Page|Yes| **Text**.|
|ZIP/Postal Code|Yes| **Text**.|

### Standard fields in a distribution list





|**Field**|**Edit in view**|**Data type and meaning**|
|:-----|:-----|:-----|
|Categories|Yes| **Text**. Field used to group and find related items. Multiple categories are separated by commas.|
|Distribution List Name|Yes| **Text**. The name of the distribution list.|
|Icon|Yes|Internal data type.|
|In Folder|No| **Text**. Name of the folder that contains the distribution list item.|
|Message Class|No|Specifies the message class for the type of item. |
|Modified|No| **Date/Time**. Last time the distribution list item was modified.|
|Outlook Internal Version|No|For administrator use only.|
|Outlook Version|No| **Text**. Version of Outlook that the distribution list item is created in.|
|Sensitivity|No|The following settings apply: **0** Normal **1** Personal **2** Private **3** Confidential|
|Size|No| **Number**. Number of bytes used by the distribution list item.|

### Standard fields in a Journal entry





|**Field**|**Edit in view**|**Data type and meaning**|
|:-----|:-----|:-----|
|Attachment|No| **Yes/No**.|
|Billing Information|Yes| **Text**.|
|Categories|Yes| **Text**. User-defined field used to group and find related items. Multiple categories are separated by commas.|
|Company|Yes| **Text**.|
|Contact|No| **Text**. Name of the contact the  **Journal** entry is recorded for.|
|Contacts|No| **Text**. Names of contacts linked to this item. Multiple names are separated by commas.|
|Created|No| **Date/Time**. Date and time the  **Journal** entry is created.|
|Do Not AutoArchive|Yes| **Yes/No**. Specifies whether to archive the  **Journal** entry.|
|Duration|No| **Duration**. Saved as minutes.|
|End|Yes| **Date/Time**. End date and time set for the  **Journal** entry.|
|Entry Type|No| **Text**. Type of entry made for the  **Journal** entry.|
|Icon|Yes|Internal data type.|
|In Folder|No| **Text**. Name of the folder that contains the  **Journal** entry.|
|Message Class|No| **Text**. Specifies the message class for the type of item.|
|Mileage|Yes| **Text**.|
|Modified|No| **Date/Time**. Last time the  **Journal** entry was modified.|
|Notes |No| **Text**. First 255 characters in the body of the  **Journal** entry.|
|Outlook Internal Version|No|For administrator use only.|
|Outlook Version|No| **Text**. Version of Outlook that the  **Journal** entry is created in.|
|Read|No| **Yes/No**. Specifies whether the  **Journal** entry has been marked as read.|
|Sensitivity|No|The following settings apply: **0** Normal **1** Personal **2** Private **3** Confidential|
|Size|No| **Number**. Number of bytes used by the  **Journal** entry.|
|Start|Yes| **Date/Time**. Start time for the  **Journal** entry.|
|Subject|Yes| **Text**.|

### Standard fields in an e-mail message





|**Field**|**Edit in view**|**Data type and meaning**|
|:-----|:-----|:-----|
|Attachment|No| **Yes/No**.|
|BCC|No| **Text**. Names in the  **Bcc** box of a message.|
|Billing Information|Yes| **Text**.|
|Categories|Yes| **Text**. Field used to group and find related items. Multiple categories are separated by commas.|
|CC|No| **Text**. Names in the  **Cc** box of a message.|
|Contacts|No| **Text**. Names of contacts linked to this item. Multiple names are separated by commas.|
|Conversation|No| **Text**. Value of the  **Subject** field in the original message in a conversation.|
|Created|No| **Date/Time**. Date and time the message is created.|
|Defer Until|Yes| **Date/Time**. Date and time a message is to be delivered. The server delays delivery of the message.|
|Do Not AutoArchive|Yes| **Yes/No**. Specifies whether to archive the message.|
|Due By|Yes| **Date/Time**. Date and time the action associated with a Message Flag is set to be completed by. When a value is entered for Due By, the Flag Status field is set to 2.|
|Expires|Yes| **Date/Time**. Date and time a message expires.|
|Flag Status|Yes|The following settings apply: **0** Normal **1** Completed **2** Flagged|
|Follow Up Flag|Yes| **Text**.|
|From|No| **Text**. Names in the  **From** box in a message.|
|Have Replies Sent To|No| **Text**. Names in the  **Have replies sent to** box in a message.|
|Header status|No|Indicates the download state of the message.|
|Icon|Yes|Internal data type.|
|Importance|Yes|The following settings apply: **0** Low importance **1** Normal importance **2** High importance|
|In Folder|No| **Text**. Name of the folder that contains the message.|
|Junk E-mail type|Yes|Internal data type.|
|Message|No| **Text**. Value of the text box in a message.|
|Message Class|No|Specifies the message class for the type of item. |
|Message Flag|Yes| **Text**. Action associated with a Message Flag. When a value is entered for Message Flag, the Flag Status field is set to 2.|
|Mileage|Yes| **Text**.|
|Modified|No| **Date/Time**. Last date and time the message was modified.|
|Outlook Internal Version|No|For administrator use only.|
|Outlook Version|No| **Text**. Version of Outlook that the message is created in.|
|Read|No| **True/False**. Specifies whether the message has been marked as read.|
|Receipt Requested|No| **Yes/No**. Indicates whether message was sent with a read or delivery receipt requested.|
|Received|No| **Date/Time**. Date and time the message is received.|
|Relevance|Yes| **Number**. User-defined significance.|
|Remote Status|No|Specifies the status of Remote Mail header. The following settings apply: **0** None **1** Marked **2** Marked for download **3** Marked for copy|
|Retrieval Time|No| **Duration**. Specifies the time it takes to download the message with Remote Mail. Saved as minutes.|
|Sensitivity|No|The following settings apply: **0** Normal **1** Personal **2** Private **3** Confidential|
|Sent|No| **Date/Time**. Date and time the message is sent.|
|Size|No| **Number**. Number of bytes used by the message.|
|Subject|Yes| **Text**.|
|To|No| **Text**. Names in the  **To** box in a message.|
|Tracking Status|No|Tracking status of a message. The following settings apply: **1** Delivered **5** Read **6** Not Read|

### Standard fields in a note





|**Field**|**Edit in view**|**Data type and meaning**|
|:-----|:-----|:-----|
|Categories|Yes| **Text**. Field used to group and find related items. Multiple categories are separated by commas.|
|Color|Yes|The following settings apply: **0** Blue **1** Green **2** Pink **3** Yellow **4** White|
|Content|No| **Text**. Value of the text box of a note.|
|Created|No| **Date/Time**. Date and time the note is created.|
|Do Not AutoArchive|Yes| **Yes/No**. Specifies whether to archive the note.|
|Icon|Yes|Internal data type.|
|In Folder|No| **Text**. Name of the folder that contains the note.|
|Message Class|No| **Text**. Specifies the message class for the type of item. |
|Modified|No| **Date/Time**. Last date and time the note was modified.|
|Outlook Internal Version|No|For administrator use only.|
|Outlook Version|No| **Text**. Version of Outlook that the note is created in.|
|Read|No| **Yes/No**. Specifies whether the note has been marked as read.|
|Size|No| **Number**. Number of bytes used by the note.|
|Subject|No| **Text**.|

### Standard fields in a task





|**Field**|**Edit in view**|**Data type and meaning**|
|:-----|:-----|:-----|
|% Complete|Yes| **Percent**.|
|Actual Work|Yes| **Duration**. Time spent on a task. Saved as minutes.|
|Assigned|No|The following settings apply: **0** Not assigned **1** Assigned by me **2** Assigned to me|
|Attachment|No| **Yes/No**.|
|Billing Information|Yes| **Text**.|
|Categories|Yes| **Text**. Field used to group and find related items. Multiple categories are separated by commas.|
|Company|Yes| **Text**.|
|Complete|Yes| **Yes/No**. Specifies whether the task is marked as completed.|
|Contacts|No| **Text**. Names of contacts linked to this item. Multiple names are separated by commas.|
|Conversation|No| **Text**. Value of the  **Subject** field of the original task.|
|Created|No| **Date/Time**. Date and time the task is created.|
|Date Completed|Yes for non-recurring tasks. No for recurring tasks.| **Date/Time**. Date and time the task is completed. If the  **Date Completed** field has a value, the **Complete** field is set to Yes, and the **% Complete** field is set to 100.|
|Do Not AutoArchive|Yes| **Yes/No**. Specifies whether to archive the task.|
|Due Date|Yes| **Date/Time**.|
|Icon|Yes|Internal data type.|
|In Folder|No| **Text**. Name of the folder that contains the task.|
|Message Class|No| **Text**. Specifies the message class of the type of item. |
|Mileage|Yes| **Text**.|
|Modified|No| **Date/Time**. Last date and time the task was modified.|
|Notes|Yes| **Text**. Value of the text box of the task.|
|Outlook Internal Version|No|For administrator use only.|
|Outlook Version|No| **Text**. Version of Outlook that the task is created in.|
|Owner|Yes, if task is in a public folder; otherwise, No.| **Text**. Owner of the task.|
|Priority|Yes|The following settings apply: **0** Low priority **1** Normal priority **2** High priority|
|Read|No| **Yes/No**. Specifies whether the task has been marked as read.|
|Recurring|No| **Yes/No**. Specifies whether the task recurs.|
|Reminder|No| **Yes/No**. Specifies whether a reminder has been set for the task.|
|Reminder Override Default|Yes| **Yes/No**.|
|Reminder Sound|Yes| **Yes/No**. Specifies whether to play the sound file as a reminder.|
|Reminder Sound File|Yes| **Text**. Path of the sound file to be played as a reminder.|
|Reminder Time|Yes| **Date/Time**. Date and time that a reminder is run for a task.|
|Request Status|No|The following settings apply: **0** No setting **1** Not responded **2** Accepted **3** Declined|
|Requested By|No| **Text**. In a task request, the person's name who assigned the task.|
|Role|Yes| **Text**.|
|Schedule+ Priority|Yes| **Text**.|
|Sensitivity|No|The following settings apply: **0** Normal **1** Personal **2** Private **3** Confidential|
|Size|No| **Number**. Number of bytes used by the task.|
|Start Date|Yes| **Date/Time**. Start date and time for the task.|
|Status|Yes|The following settings apply: **0** Not Started **1** In Progress **2** Completed **3** Waiting on someone else **4** Deferred|
|Subject|Yes| **Text**.|
|Team Task|Yes| **Yes/No**.|
|To|No|The owner of an assigned task.|
|Total Work|Yes| **Duration**. Time the task is expected to take. Saved as minutes.|

### Standard fields in a post





|**Field**|**Edit in view**|**Data type and meaning**|
|:-----|:-----|:-----|
|Attachment|No| **Yes/No**.|
|Billing Information|Yes| **Text**.|
|Categories|Yes| **Text**. Field used to group and find related items. Multiple categories are separated by commas.|
|Conversation|No| **Text**. Value of the  **Subject** field in the original item in a conversation.|
|Created|No| **Date/Time**. Date and time the message is created.|
|Defer Until|Yes| **Date/Time**. Date and time a message is to be delivered. The server delays delivery of the message.|
|Do Not AutoArchive|Yes| **Yes/No**. Specifies whether to archive the message.|
|Expires|Yes| **Date/Time**. Date and time a message expires.|
|From|No| **Text**. Names in the  **From** box in a message.|
|Header status|No|Indicates the download state of the message.|
|Icon|Yes|Internal data type.|
|Importance|Yes|The following settings apply: **0** Low importance **1** Normal importance **2** High importance|
|In Folder|No| **Text**. Name of the folder that contains the message.|
|Message|No| **Text**. Value of the text box in a message.|
|Message Class|No|Specifies the message class for the type of item. |
|Mileage|Yes| **Text**.|
|Modified|No| **Date/Time**. Last date and time the message was modified.|
|Outlook Internal Version|No|For administrator use only.|
|Outlook Version|No| **Text**. Version of Outlook that the message is created in.|
|Read|No| **True/False**. Specifies whether the message has been marked as read.|
|Received|No| **Date/Time**. Date and time the message is received.|
|Remote Status|No|Specifies the status of Remote Mail header. The following settings apply: **0** None **1** Marked **2** Marked for download **3** Marked for copy|
|Retrieval Time|No| **Duration**. Specifies the time it takes to download the message with Remote Mail. Saved as minutes.|
|Sensitivity|No|The following settings apply: **0** Normal **1** Personal **2** Private **3** Confidential|
|Sent|No| **Date/Time**. Date and time the message is sent.|
|Size|No| **Number**. Number of bytes used by the message.|
|Subject|Yes| **Text**.|

