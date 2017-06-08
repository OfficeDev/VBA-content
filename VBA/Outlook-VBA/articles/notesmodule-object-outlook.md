---
title: NotesModule Object (Outlook)
keywords: vbaol11.chm3198
f1_keywords:
- vbaol11.chm3198
ms.prod: outlook
api_name:
- Outlook.NotesModule
ms.assetid: cdbdde08-0773-a78d-3809-a3811975bcc1
ms.date: 06/08/2017
---


# NotesModule Object (Outlook)

Represents the  **Notes** navigation module in the Navigation Pane of an explorer.


## Remarks

The  **NotesModule** object, derived from the **[NavigationModule](navigationmodule-object-outlook.md)** object, provides access to the navigation groups contained in the **Notes** navigation module of the Navigation Pane for an explorer. Use the **[GetNavigationModule](navigationmodules-getnavigationmodule-method-outlook.md)** method or the **[Item](navigationmodules-item-method-outlook.md)** method of the **[NavigationModules](navigationmodules-object-outlook.md)** collection for the parent **[NavigationPane](navigationpane-object-outlook.md)** object to retrieve a **NavigationModule** object, then use the **[NavigationModuleType](navigationmodule-navigationmoduletype-property-outlook.md)** property of the **NavigationModule** object to retrieve the navigation module type. If the **NavigationModuleType** property is set to **olModuleNotes**, you can then cast the **Module** object reference as a **NotesModule** object to access the **[NavigationGroups](notesmodule-navigationgroups-property-outlook.md)** property for that navigation module.

You can use the  **[Visible](notesmodule-visible-property-outlook.md)** property to determine if the navigation module is visible and the **[Position](notesmodule-position-property-outlook.md)** property to return or set the display position of the navigation module within the Navigation Pane. You can use the **[Name](notesmodule-name-property-outlook.md)** property to return the display name of the **Notes** navigation module within the Navigation Pane.


## Properties



|**Name**|
|:-----|
|[Application](notesmodule-application-property-outlook.md)|
|[Class](notesmodule-class-property-outlook.md)|
|[Name](notesmodule-name-property-outlook.md)|
|[NavigationGroups](notesmodule-navigationgroups-property-outlook.md)|
|[NavigationModuleType](notesmodule-navigationmoduletype-property-outlook.md)|
|[Parent](notesmodule-parent-property-outlook.md)|
|[Position](notesmodule-position-property-outlook.md)|
|[Session](notesmodule-session-property-outlook.md)|
|[Visible](notesmodule-visible-property-outlook.md)|

## See also


#### Other resources


[Outlook Object Model Reference](http://msdn.microsoft.com/library/73221b13-d8d8-99b8-3394-b95dbbfd5ddc%28Office.15%29.aspx)
