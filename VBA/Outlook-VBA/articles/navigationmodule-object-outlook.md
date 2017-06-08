---
title: NavigationModule Object (Outlook)
keywords: vbaol11.chm3211
f1_keywords:
- vbaol11.chm3211
ms.prod: outlook
api_name:
- Outlook.NavigationModule
ms.assetid: 76565eaf-1e64-f5d4-b90f-ba156863802c
ms.date: 06/08/2017
---


# NavigationModule Object (Outlook)

Represents a navigation module in the Navigation Pane.


## Remarks

The  **NavigationModule** object provides access to the various navigation modules that are displayed in the Microsoft Outlook Navigation Pane. The following objects are derived from the **NavigationModule** object:


-  **[CalendarModule](calendarmodule-object-outlook.md)**
    
-  **[ContactsModule](contactsmodule-object-outlook.md)**
    
-  **[JournalModule](journalmodule-object-outlook.md)**
    
-  **[MailModule](mailmodule-object-outlook.md)**
    
-  **[NotesModule](notesmodule-object-outlook.md)**
    
-  **[TasksModule](tasksmodule-object-outlook.md)**
    
-  **[SolutionsModule](solutionsmodule-object-outlook.md)**
    
 Use the **[GetNavigationModule](navigationmodules-getnavigationmodule-method-outlook.md)** method or the **[Item](navigationmodules-item-method-outlook.md)** method of the **[NavigationModules](navigationmodules-object-outlook.md)** collection for the parent **[NavigationPane](navigationpane-object-outlook.md)** object to retrieve a **NavigationModule** object, then use the **[NavigationModuleType](navigationmodule-navigationmoduletype-property-outlook.md)** property of the **NavigationModule** object to retrieve the module type. Depending on the value of the **NavigationModuleType** property, you can then cast the **NavigationModule** object reference as one of the objects listed in the previous paragraph to access the **[NavigationGroups](calendarmodule-navigationgroups-property-outlook.md)** property for that object, such as a **CalendarModule** object.

The  **Shortcuts** and **Folder List** navigation modules do not have a corresponding object, such as **MailModule**, because they do not support programmatic access to navigation groups or navigation folders. You can use the **NavigationModule** object to access the properties of the **Shortcuts** and **Folder List** modules.

You can use the  **[Visible](navigationmodule-visible-property-outlook.md)** property to determine whether the navigation module is visible, and use the **[Position](navigationmodule-position-property-outlook.md)** property to return or set the display position of the navigation module within the Navigation Pane. You can also use the **[Name](navigationmodule-name-property-outlook.md)** property to return the display name of the navigation module in the Navigation Pane.


## Properties



|**Name**|
|:-----|
|[Application](navigationmodule-application-property-outlook.md)|
|[Class](navigationmodule-class-property-outlook.md)|
|[Name](navigationmodule-name-property-outlook.md)|
|[NavigationModuleType](navigationmodule-navigationmoduletype-property-outlook.md)|
|[Parent](navigationmodule-parent-property-outlook.md)|
|[Position](navigationmodule-position-property-outlook.md)|
|[Session](navigationmodule-session-property-outlook.md)|
|[Visible](navigationmodule-visible-property-outlook.md)|

## See also


#### Other resources


[Outlook Object Model Reference](http://msdn.microsoft.com/library/73221b13-d8d8-99b8-3394-b95dbbfd5ddc%28Office.15%29.aspx)
