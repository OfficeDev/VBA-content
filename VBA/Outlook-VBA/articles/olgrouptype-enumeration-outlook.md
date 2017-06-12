---
title: OlGroupType Enumeration (Outlook)
keywords: vbaol11.chm3144
f1_keywords:
- vbaol11.chm3144
ms.prod: outlook
api_name:
- Outlook.OlGroupType
ms.assetid: 2a5ee820-41fa-91fc-2ce0-46d97fc4bf11
ms.date: 06/08/2017
---


# OlGroupType Enumeration (Outlook)

Identifies the group type of a  **[NavigationGroup](navigationgroup-object-outlook.md)** object.



|**Name**|**Value**|**Description**|
|:-----|:-----|:-----|
| **olCustomFoldersGroup**|0|Identifies a user-defined navigation group, added using either the Outlook user interface or an add-in.|
| **olFavoriteFoldersGroup**|4|Identifies the  **Favorite Folders** navigation group. This navigation group exists only within the **[NavigationGroups](mailmodule-navigationgroups-property-outlook.md)** collection of a **[MailModule](mailmodule-object-outlook.md)** object and cannot be created in or accessed from other modules.|
| **olMyFoldersGroup**|1|Identifies a navigation group that, by default, contains any folders that are part of the local store.|
| **olOtherFoldersGroup**|3|Identifies a navigation group that, by default, contains shared folders from sources other than that of other persons.|
| **olPeopleFoldersGroup**|2|Identifies a navigation group that, by default, contains shared folders from other persons.|
| **olReadOnlyGroup**|6|Identifies a navigation group that is, by default, read-only and no folders can be added or removed from that navigation group. This does not imply the folders themselves are read-only, and write access to the folders depends on how the folders are set up.|
| **olRoomsGroup**|5|Identifies the  **Rooms** navigation group in the **Calendar** navigation module.|

