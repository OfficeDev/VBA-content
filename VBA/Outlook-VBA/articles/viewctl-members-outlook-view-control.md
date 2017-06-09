---
title: ViewCtl Members (Outlook View Control)
ms.prod: outlook
ms.assetid: 32df30fd-d02c-30c4-7474-0dc359f99f46
ms.date: 06/08/2017
---


# ViewCtl Members (Outlook View Control)

The Microsoft Outlook View Control displays information about a specific folder and can be integrated into solutions that provide access to Outlook data. The  **ViewCtl** object provides programmatic access to the View Control. The control can be placed in any container that supports ActiveX® controls, including an HTML page that is hosted in Outlook as a Folder Home Page, or a custom Outlook form. If the View Control is placed in an HTML page that is hosted in a browser such as Internet Explorer, some functions of the control are disabled for security.


## Methods



|**Name**|**Description**|
|:-----|:-----|
| **[AddressBook](viewctl-addressbook-method-outlook-view-control.md)**|Displays the Microsoft Outlook  **Address Book** dialog box.|
| **[AddToPFFavorites](viewctl-addtopffavorites-method-outlook-view-control.md)**|Adds the current public folder to the user's Microsoft Exchange Server  **Favorites** public folder.|
| **[AdvancedFind](viewctl-advancedfind-method-outlook-view-control.md)**|Displays the Microsoft Outlook  **Advanced Find** dialog box.|
| **[Categories](viewctl-categories-method-outlook-view-control.md)**|Displays the Microsoft Outlook  **Categories** dialog box for the currently selected item or items in the control.|
| **[CollapseAllGroups](viewctl-collapseallgroups-method-outlook-view-control.md)**|Collapses (closes) all groups that are displayed in the control.|
| **[CollapseGroup](viewctl-collapsegroup-method-outlook-view-control.md)**|Collapses (closes) the group that is currently selected in the control. |
| **[CustomizeView](viewctl-customizeview-method-outlook-view-control.md)**|Displays the Microsoft Outlook  **View Summary** dialog box.|
| **[Delete](viewctl-delete-method-outlook-view-control.md)**|After prompting the user to confirm, deletes the groups or items that are currently selected in the control. |
| **[ExpandAllGroups](viewctl-expandallgroups-method-outlook-view-control.md)**|Expands (opens) all groups that are displayed in the control. |
| **[ExpandGroup](viewctl-expandgroup-method-outlook-view-control.md)**|Expands (opens) the group that is currently selected in the control. |
| **[FlagItem](viewctl-flagitem-method-outlook-view-control.md)**|Displays the Microsoft Outlook  **Flag for Follow Up** dialog box for the selected item.|
| **[ForceUpdate](viewctl-forceupdate-method-outlook-view-control.md)**|Refreshes the view in the control, applying any property changes made since the  **[DeferUpdate](viewctl-deferupdate-property-outlook-view-control.md)** property was set to **True**.|
| **[Forward](viewctl-forward-method-outlook-view-control.md)**|Executes the Forward action for the item or items that are selected in the control.|
| **[GoToDate](viewctl-gotodate-method-outlook-view-control.md)**|Opens a calendar view of a specific date.|
| **[NewAppointment](viewctl-newappointment-method-outlook-view-control.md)**|Creates and displays a new appointment.|
| **[NewContact](viewctl-newcontact-method-outlook-view-control.md)**|Creates and displays a new contact.|
| **[NewDefaultItem](viewctl-newdefaultitem-method-outlook-view-control.md)**|Creates and displays a new Microsoft Outlook item. |
| **[NewDistributionList](viewctl-newdistributionlist-method-outlook-view-control.md)**|Creates and displays a new distribution list.|
| **[NewForm](viewctl-newform-method-outlook-view-control.md)**|Displays the Microsoft Outlook  **Choose Form** dialog box.|
| **[NewJournalEntry](viewctl-newjournalentry-method-outlook-view-control.md)**|Creates and displays a new journal entry.|
| **[NewMeetingRequest](viewctl-newmeetingrequest-method-outlook-view-control.md)**|Creates and displays a new meeting request.|
| **[NewMessage](viewctl-newmessage-method-outlook-view-control.md)**|Creates and displays a new e-mail message.|
| **[NewNote](viewctl-newnote-method-outlook-view-control.md)**|Creates and displays a new note item.|
| **[NewPost](viewctl-newpost-method-outlook-view-control.md)**|Creates and displays a new post item.|
| **[NewTask](viewctl-newtask-method-outlook-view-control.md)**|Creates and displays a new task.|
| **[NewTaskRequest](viewctl-newtaskrequest-method-outlook-view-control.md)**|Creates and displays a new task request.|
| **[Open](viewctl-open-method-outlook-view-control.md)**|Opens the item or items that are currently selected in the control.|
| **[OpenSharedDefaultFolder](viewctl-openshareddefaultfolder-method-outlook-view-control.md)**|Displays a specified user's default folder in the control.|
| **[PrintItem](viewctl-printitem-method-outlook-view-control.md)**|Prints the items that are currently selected in the control. |
| **[Reply](viewctl-reply-method-outlook-view-control.md)**|Executes the Reply action for the item or items selected in the control.|
| **[ReplyAll](viewctl-replyall-method-outlook-view-control.md)**|Executes the ReplyAll action for the item or items that are selected in the control.|
| **[ReplyInFolder](viewctl-replyinfolder-method-outlook-view-control.md)**|Creates a post item for each message that is currently selected in the control.|
| **[SaveAs](viewctl-saveas-method-outlook-view-control.md)**|Saves the items that are selected in the control as a single file.|
| **[SendAndReceive](viewctl-sendandreceive-method-outlook-view-control.md)**|Sends all messages that are in the  **Outbox** folder and checks for new messages.|
| **[ShowFields](viewctl-showfields-method-outlook-view-control.md)**|Displays the Microsoft Outlook  **Show Fields** dialog box.|
| **[Sort](viewctl-sort-method-outlook-view-control.md)**|Displays the Microsoft Outlook  **Sort** dialog box.|
| **[SynchFolder](viewctl-synchfolder-method-outlook-view-control.md)**|Synchronizes the online and offline folders that are displayed in the control. |

## Properties



|**Name**|**Description**|
|:-----|:-----|
| **[ActiveFolder](viewctl-activefolder-property-outlook-view-control.md)**|Returns an object that represents the folder displayed in the control. Read-only.|
| **[DeferUpdate](viewctl-deferupdate-property-outlook-view-control.md)**|Gets or sets a  **Boolean** value that indicates whether property changes affect the control display. Read/write.|
| **[EnableRowPersistance](viewctl-enablerowpersistance-property-outlook-view-control.md)**|Gets or sets a value that indicates whether the View Control retains state information about the last selected row. Read/write.|
| **[Filter](viewctl-filter-property-outlook-view-control.md)**|Gets or sets a  **String** that represents the Distributed Authoring and Versioning (DAV) Searching and Locating (DASL) statement used to restrict the display to a specified subset of data. Read/write.|
| **[FilterAppend](viewctl-filterappend-property-outlook-view-control.md)**|Gets or sets a  **String** that represents the additional criteria to add to the filter settings. Read/write.|
| **[Folder](viewctl-folder-property-outlook-view-control.md)**|Gets or sets a  **String** that represents the path of the folder displayed by the control.|
| **[ItemCount](viewctl-itemcount-property-outlook-view-control.md)**|Returns a  **Long** that indicates the count of objects in the current folder displayed in the control. Read-only.|
| **[Namespace](viewctl-namespace-property-outlook-view-control.md)**|Returns or sets a  **String** that represents the namespace property of the control. Read/write.|
| **[OutlookApplication](viewctl-outlookapplication-property-outlook-view-control.md)**|Returns an object that represents the container object for the control. Read-only.|
| **[Restriction](viewctl-restriction-property-outlook-view-control.md)**|Sets or returns a  **String** that represents a filter to the items that are displayed in the control. As a result, the control displays only those items that match the filter. Read/write.|
| **[SelectedDate](viewctl-selecteddate-property-outlook-view-control.md)**|Returns or sets the selected date. Read-only.|
| **[Selection](viewctl-selection-property-outlook-view-control.md)**|Returns a  **[Selection](selection-object-outlook.md)** object that consists of one or more items that are selected in the current view. Read-only.|
| **[View](viewctl-view-property-outlook-view-control.md)**|Returns or sets a  **String** that represents the name of the view in the control. Read/write.|
| **[ViewXML](viewctl-viewxml-property-outlook-view-control.md)**|Returns or sets a  **String** that represents the view implementation via XML. Read/write.|

## Events



|**Name**|**Description**|
|:-----|:-----|
| **[Activate](viewctl-activate-event-outlook-view-control.md)**|Occurs when a View Control becomes the active element on the page, either as a result of user action or through program code.|
| **[BeforeViewSwitch](viewctl-beforeviewswitch-event-outlook-view-control.md)**|Occurs before Microsoft Outlook changes the view that is applied to the folder displayed in the View Control element, either as a result of user action or through program code. |
| **[SelectionChange](viewctl-selectionchange-event-outlook-view-control.md)**|Occurs when the selection of the current view changes. |
| **[ViewSwitch](viewctl-viewswitch-event-outlook-view-control.md)**|Occurs when Microsoft Outlook changes the view that is applied to the folder displayed in the View Control element, either as a result of user action or through program code.|

