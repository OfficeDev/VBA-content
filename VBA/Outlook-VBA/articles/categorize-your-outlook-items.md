---
title: Categorize Your Outlook Items
ms.prod: outlook
ms.assetid: e8cfb450-b8b0-bee6-fdf0-d0a92bf9af56
ms.date: 06/08/2017
---


# Categorize Your Outlook Items

Microsoft Outlook provides color categorization functionality, in which Outlook items can be categorized and displayed by category. Multiple color categories can be applied to a single Outlook item, and Outlook items can be grouped or sorted by color category. Shortcut keys can be assigned to each color category, to allow users to more easily categorize items. Color categories are user-defined, and can be created, deleted, and changed either programmatically or by user action within the Outlook user interface.

The  **[Category](category-object-outlook.md)** object represents a single user-defined color category in the Master Category List, the list of color categories presented in the Outlook user interface and represented by the **[Categories](namespace-categories-property-outlook.md)** collection of the **[NameSpace](namespace-object-outlook.md)** object. **Category** objects are identified with a globally unique identifier (GUID) when created, and this identifier cannot be changed. However, the name, color, and shortcut key associated with a color category can be changed by setting the **[Name](category-name-property-outlook.md)**,  **[Color](category-color-property-outlook.md)**, and  **[ShortcutKey](category-shortcutkey-property-outlook.md)** properties, respectively, of the **Category** object. The **[CategoryID](category-categoryid-property-outlook.md)** property can be used to retrieve the identifier of a **Category** object.

## Assigning Categories to Outlook Items

Categories can be assigned to Outlook items by specifying the names of the appropriate  **Category** objects in a comma-delimited string in the **Categories** property of the following objects:



| <strong><a href="appointmentitem-object-outlook.md" data-raw-source="[AppointmentItem](appointmentitem-object-outlook.md)">AppointmentItem</a></strong>                                                                                                         | <strong><a href="remoteitem-object-outlook.md" data-raw-source="[RemoteItem](remoteitem-object-outlook.md)">RemoteItem</a></strong>                                                 |
|:----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------|:------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------|
| <strong><a href="contactitem-object-outlook.md" data-raw-source="[ContactItem](contactitem-object-outlook.md)">ContactItem</a></strong>                                                                                                                         | <strong><a href="reportitem-object-outlook.md" data-raw-source="[ReportItem](reportitem-object-outlook.md)">ReportItem</a></strong>                                                 |
| <strong><a href="distlistitem-object-outlook.md" data-raw-source="[DistListItem](distlistitem-object-outlook.md)">DistListItem</a></strong>                                                                                                                     | <strong><a href="sharingitem-object-outlook.md" data-raw-source="[SharingItem](sharingitem-object-outlook.md)">SharingItem</a></strong>                                             |
| <strong><a href="documentitem-object-outlook.md" data-raw-source="[DocumentItem](documentitem-object-outlook.md)">DocumentItem</a></strong>                                                                                                                     | <strong><a href="postitem-object-outlook.md" data-raw-source="[PostItem](postitem-object-outlook.md)">PostItem</a></strong>                                                         |
| <strong><a href="journalitem-object-outlook.md" data-raw-source="[JournalItem](journalitem-object-outlook.md)">JournalItem</a></strong>                                                                                                                         | <strong><a href="taskitem-object-outlook.md" data-raw-source="[TaskItem](taskitem-object-outlook.md)">TaskItem</a></strong>                                                         |
| <strong><a href="mailitem-object-outlook.md" data-raw-source="[MailItem](mailitem-object-outlook.md)">MailItem</a></strong>                                                                                                                                     | <strong><a href="taskrequestacceptitem-object-outlook.md" data-raw-source="[TaskRequestAcceptItem](taskrequestacceptitem-object-outlook.md)">TaskRequestAcceptItem</a></strong>     |
| <strong><a href="meetingitem-object-outlook.md" data-raw-source="[MeetingItem](meetingitem-object-outlook.md)">MeetingItem</a></strong>                                                                                                                         | <strong><a href="taskrequestdeclineitem-object-outlook.md" data-raw-source="[TaskRequestDeclineItem](taskrequestdeclineitem-object-outlook.md)">TaskRequestDeclineItem</a></strong> |
| <strong><a href="http://msdn.microsoft.com/library/da8149d5-66d3-ea02-941f-e7f2f9eb6bc3%28Office.15%29.aspx" data-raw-source="[MobileItem](http://msdn.microsoft.com/library/da8149d5-66d3-ea02-941f-e7f2f9eb6bc3%28Office.15%29.aspx)">MobileItem</a></strong> | <strong><a href="taskrequestitem-object-outlook.md" data-raw-source="[TaskRequestItem](taskrequestitem-object-outlook.md)">TaskRequestItem</a></strong>                             |
| <strong><a href="noteitem-object-outlook.md" data-raw-source="[NoteItem](noteitem-object-outlook.md)">NoteItem</a></strong>                                                                                                                                     | <strong><a href="taskrequestupdateitem-object-outlook.md" data-raw-source="[TaskRequestUpdateItem](taskrequestupdateitem-object-outlook.md)">TaskRequestUpdateItem</a></strong>     |

Outlook items are displayed based on the category name stored in the  **Categories** property of that Outlook item. Because category names are stored as part of the Outlook item, it is possible to have a category name in an Outlook item that is not present in the Master Category List. For example, a category may have been removed.

If a  **Category** object with a corresponding **Name** property value does not exist in the **Categories** collection of the **NameSpace** object that contains the Outlook item, the category name associated with that Outlook item is still displayed, but without an associated color.


