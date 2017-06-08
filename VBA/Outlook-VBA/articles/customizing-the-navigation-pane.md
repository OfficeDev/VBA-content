---
title: Customizing the Navigation Pane
ms.prod: outlook
ms.assetid: 426c3d1c-13b5-cac5-702d-87dfe71f2478
ms.date: 06/08/2017
---


# Customizing the Navigation Pane

The Navigation Pane provides access to information that pertains to the active explorer, including different views and different ways to accomplish tasks in that explorer. The  **[NavigationPane](navigationpane-object-outlook.md)** object represents the Navigation Pane for an explorer; to obtain one, call the **[NavigationPane](explorer-navigationpane-property-outlook.md)** property of the **[Explorer](explorer-object-outlook.md)** object. If the explorer does not contain a Navigation Pane, this property returns **Null** ( **Nothing** in Visual Basic).


## Navigation Modules

The Navigation Pane contains the set of navigation modules that are available in Outlook; for example, the  **Mail** module. Each navigation module is represented by a **[NavigationModule](navigationmodule-object-outlook.md)** object or by an object that is derived from the **NavigationModule** object. The **[Modules](navigationpane-modules-property-outlook.md)** property of the **NavigationPane** object provides access to the navigation modules that are in the Navigation Pane. You can use the following objects to access the corresponding navigation module:



|**Navigation module**|**Object**|
|:-----|:-----|
| **Calendar**| **[CalendarModule](calendarmodule-object-outlook.md)**|
| **Contacts**| **[ContactsModule](contactsmodule-object-outlook.md)**|
| **Journal**| **[JournalModule](journalmodule-object-outlook.md)**|
| **Folder List**| **NavigationModule**|
| **Mail**| **[MailModule](mailmodule-object-outlook.md)**|
| **Notes**| **[NotesModule](notesmodule-object-outlook.md)**|
| **Shortcuts**| **NavigationModule**|
| **Solutions**| **[SolutionsModule](solutionsmodule-object-outlook.md)**|
| **Tasks**| **[TasksModule](tasksmodule-object-outlook.md)**|
Note that the  **Solutions** module is not displayed in the Navigation Pane by default, and can only be created programmatically. The default name of the module is **Solutions**, but you can customize that name.


## Navigation Groups and Navigation Folders

Each navigation module contains a set of navigation groups. A navigation group, represented by the  **[NavigationGroup](navigationgroup-object-outlook.md)** object, is a container for navigation folders. A navigation folder, represented by the **[NavigationFolder](navigationfolder-object-outlook.md)** object, provides an access point in the Navigation Pane for a **[Folder](folder-object-outlook.md)** object. You can obtain a **NavigationGroup** object reference by using the **[NavigationGroups](navigationgroups-object-outlook.md)** property of a **CalendarModule**,  **ContactsModule**,  **JournalModule**,  **MailModule**,  **NotesModule**, or  **TasksModule** object. The **Folder List**,  **Shortcuts**, and  **Solutions** navigation modules do not contain navigation groups.

You can create and delete custom navigation groups by using the  **[NavigationGroups.Create](navigationgroups-create-method-outlook.md)** and **[NavigationGroups.Delete](navigationgroups-delete-method-outlook.md)** methods. You can identify a custom navigation group by using the **[NavigationGroup.GroupType](navigationgroup-grouptype-property-outlook.md)** property to retrieve the navigation group type for the object, and you can retrieve the default navigation group for a specified group type by using the **[NavigationGroups.GetDefaultNavigationGroup](navigationgroups-getdefaultnavigationgroup-method-outlook.md)** method.

Once you have a  **NavigationGroup** object, you can obtain a **NavigationFolder** object reference by using the **[NavigationGroup.NavigationFolders](navigationgroup-navigationfolders-property-outlook.md)** property. Each **NavigationFolder** represents a navigation folder associated with a **Folder** object. You can add navigation folders to a navigation group by using the **[NavigationFolders.Add](navigationfolders-add-method-outlook.md)** method. Only one **NavigationFolder** object can be associated with a **Folder** object at any given time, so adding a **NavigationFolder** that is associated with a given **Folder** object to a navigation group automatically removes any existing **NavigationFolder** references that are associated with that **Folder** object. You can also delete navigation folders from a navigation group by using the **[NavigationFolders.Remove](navigationfolders-remove-method-outlook.md)** method, but only if the **[NavigationFolders.IsRemovable](navigationfolder-isremovable-property-outlook.md)** property is set to **True** for the **NavigationFolder** object to be removed. You cannot remove standard navigation folders, such as the **Inbox** folder, that are defined by Outlook.


 **Note**  Navigation folders can be freely added or removed from the  **Favorite Folders** navigation group, a special navigation group that is contained by the **MailModule** object, regardless of the **IsRemovable** property value of the navigation folder.


## Displaying the Navigation Pane

The Navigation Pane can display navigation modules in either normal or collapsed mode. The  **[Visible](navigationmodule-visible-property-outlook.md)** property of a **NavigationModule** object determines whether the navigation module is displayed in the Navigation Pane; the order that the visible navigation modules are displayed is determined by the **[Position](navigationmodule-position-property-outlook.md)** property of each **NavigationModule** object.

You can use the  **[IsCollapsed](navigationpane-iscollapsed-property-outlook.md)** property to determine which mode the **NavigationPane** object uses. In normal mode, the visible navigation modules in the Navigation Pane are displayed as a combination of large and small buttons. The number of large buttons that are displayed in normal mode is determined by the **[DisplayedModuleCount](navigationpane-displayedmodulecount-property-outlook.md)** property. If there are more visible navigation modules than are specified by this property, the remaining visible navigation modules are displayed as small buttons at the bottom of the Navigation Pane. In collapsed mode, the visible navigation modules in the Navigation Pane are displayed as small buttons. The number of small buttons displayed in collapsed mode is determined by the **DisplayedModuleCount** property. If there are more visible navigation modules than are specified by this property, the remaining visible navigation modules are not displayed.

To change the current navigation module, set the  **[CurrentModule](navigationpane-currentmodule-property-outlook.md)** property of the **NavigationPane** object to one of the **NavigationModule** objects in the navigation pane.

In each navigation module, the  **[NavigationGroup.Position](navigationgroup-position-property-outlook.md)** property determines the display order of the navigation groups. Similarly, the **[NavigationFolder.Position](navigationfolder-position-property-outlook.md)** property determines the display order of navigation folders within each navigation group. If a **NavigationFolder** object represents a calendar folder, the **[IsSideBySide](navigationfolder-issidebyside-property-outlook.md)** property determines if the contents of the calendar folder are displayed in side-by-side or overlay mode.


## Handling Navigation Pane Events

The  **NavigationPane** object provides the **[ModuleSwitch](navigationpane-moduleswitch-event-outlook.md)** event so that add-ins can identify when the current navigation module changes in the Navigation Pane, either programmatically or by user action.

The  **NavigationGroups** object provides the **[NavigationFolderAdd](navigationgroups-navigationfolderadd-event-outlook.md)** and **[NavigationFolderRemove](navigationgroups-navigationfolderremove-event-outlook.md)** events so that add-ins can identify when a navigation folder is added or removed from a **NavigationGroup** object in the collection. The **NavigationGroups** object also provides the **[SelectedChange](navigationgroups-selectedchange-event-outlook.md)** event. Add-ins use that event to identify when the **[IsSelected](navigationfolder-isselected-property-outlook.md)** property of a navigation folder that is associated with a calendar folder changes in the Navigation Pane, either programmatically or by user action.

To detect a user change a folder in the Folder List, use the  **[BeforeFolderSwitch](explorer-beforefolderswitch-event-outlook.md)** and **[FolderSwitch](explorer-folderswitch-event-outlook.md)** events of the **[Explorer](explorer-object-outlook.md)** object. Similarly, to detect when the **Solutions** module is first displayed in the Navigation Pane, or to detect a user click a different folder in the **Solutions** module, use the **BeforeFolderSwitch** and **FolderSwitch** events.


## See also


#### Concepts


 [Adding Solution-Specific Folders to the Solutions Module](adding-solution-specific-folders-to-the-solutions-module.md)

