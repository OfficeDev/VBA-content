---
title: OlDefaultSelectNamesDisplayMode Enumeration (Outlook)
keywords: vbaol11.chm3107
f1_keywords:
- vbaol11.chm3107
ms.prod: outlook
api_name:
- Outlook.OlDefaultSelectNamesDisplayMode
ms.assetid: 4a9c2183-c704-fc4d-e3c8-32c53b9688bb
ms.date: 06/08/2017
---


# OlDefaultSelectNamesDisplayMode Enumeration (Outlook)

Specifies the default caption, the number of buttons, the button labels, and the address lists to display in the  **Select Names** dialog box without requiring a resource file for any localized caption and labels.



|**Name**|**Value**|**Description**|
|:-----|:-----|:-----|
| **olDefaultDelegates**|6|Displays one edit box for To recipients, uses localized string representing "Add" for the To button, and localized string representing "Add Users" for the caption.  **CcLabel** and **BccLabel** are set to an empty string. Sets **[SelectNamesDialog.AllowMultipleSelection](selectnamesdialog-allowmultipleselection-property-outlook.md)** to **True** and **[SelectNamesDialog.NumberOfRecipientSelectors](selectnamesdialog-numberofrecipientselectors-property-outlook.md)** to **olTo**.|
| **olDefaultMail**|1|Displays three edit boxes for To, Cc, and Bcc recipients, uses localized strings representing "To", "Cc", and "Bcc" for To, Cc, and Bcc buttons, and localized string representing "Select Names" for the caption. Sets  **AllowMultipleSelection** to **True** and **NumberOfRecipientSelectors** to **olToCcBcc**.|
| **olDefaultMeeting**|2|Displays three edit boxes for Required, Optional, and Resource recipients, uses localized strings representing "Required", "Optional", and "Resources" for the To, Cc, and Bcc buttons, and localized string representing "Select Attendees and Resources" for the caption. Sets  **AllowMultipleSelection** to **True** and **NumberOfRecipientSelectors** to **olToCcBcc**.|
| **olDefaultMembers**|5|Displays one edit box for To recipients, uses localized string representing "To" for the To button, and localized string representing "Select Members" for caption.  **CcLabel** and **BccLabel** are set to an empty string. Sets **AllowMultipleSelection** to **True** and **NumberOfRecipientSelectors** to **olTo**.|
| **olDefaultPickRooms**|8|Displays one edit box for Resource recipients, uses localized string representing "Rooms" for To button, and localized string representing "Select Rooms" for caption.  **CcLabel** and **BccLabel** are set to an empty string. Sets **AllowMultipleSelection** to **True** and **NumberOfRecipientSelectors** to ****olShowTo.  **InitialDisplayList** is set to the Global Address List.|
| **olDefaultSharingRequest**|4|Displays one edit box for To recipients, uses localized string representing "To" for To button, and localized string representing "Select Names" for caption.  **CcLabel** and **BccLabel** are set to an empty string. Sets **AllowMultipleSelection** to **True** and **NumberOfRecipientSelectors** to **olTo**.|
| **olDefaultSingleName**|7|Displays no edit boxes for recipients, uses localized string representing "Select Name" for caption.  **ToLabel**,  **CcLabel**, and  **Bcclabel** are set to an empty string. Sets **AllowMultipleSelection** to **False** and **NumberOfRecipientSelectors** to **olNone**. |
| **olDefaultTask**|3|Displays one edit box for To recipients, uses localized string representing "To" for To button, and localized string representing "Select Task Recipient" for caption.  **CcLabel** and **BccLabel** are set to an empty string. Sets **AllowMultipleSelection** to **True** and **NumberOfRecipientSelectors** to **olTo**.|

