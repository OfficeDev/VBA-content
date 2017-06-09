---
title: TasksModule Object (Outlook)
keywords: vbaol11.chm3196
f1_keywords:
- vbaol11.chm3196
ms.prod: outlook
api_name:
- Outlook.TasksModule
ms.assetid: fc6ae6c9-6b13-b5f2-9506-c3dbbe709df6
ms.date: 06/08/2017
---


# TasksModule Object (Outlook)

Represents the  **Tasks** navigation module in the Navigation Pane of an explorer.


## Remarks

The  **TasksModule** object, derived from the **[NavigationModule](navigationmodule-object-outlook.md)** object, provides access to the navigation groups contained in the **Tasks** navigation module of the Navigation Pane for an explorer. Use the **[GetNavigationModule](navigationmodules-getnavigationmodule-method-outlook.md)** method or the **[Item](navigationmodules-item-method-outlook.md)** method of the **[NavigationModules](navigationmodules-object-outlook.md)** collection for the parent **[NavigationPane](navigationpane-object-outlook.md)** object to retrieve a **NavigationModule** object, then use the **[NavigationModuleType](navigationmodule-navigationmoduletype-property-outlook.md)** property of the **NavigationModule** object to retrieve the navigation module type. If the **NavigationModuleType** property is set to **olModuleTasks**, you can then cast the **NavigationModule** object reference as a **TasksModule** object to access the **[NavigationGroups](tasksmodule-navigationgroups-property-outlook.md)** property for that navigation module.

You can use the  **[Visible](tasksmodule-visible-property-outlook.md)** property to determine if the navigation module is visible and the **[Position](tasksmodule-position-property-outlook.md)** property to return or set the display position of the navigation module within the Navigation Pane. You can use the **[Name](tasksmodule-name-property-outlook.md)** property to return the display name of the **Tasks** navigation module within the Navigation Pane.


## Properties



|**Name**|
|:-----|
|[Application](tasksmodule-application-property-outlook.md)|
|[Class](tasksmodule-class-property-outlook.md)|
|[Name](tasksmodule-name-property-outlook.md)|
|[NavigationGroups](tasksmodule-navigationgroups-property-outlook.md)|
|[NavigationModuleType](tasksmodule-navigationmoduletype-property-outlook.md)|
|[Parent](tasksmodule-parent-property-outlook.md)|
|[Position](tasksmodule-position-property-outlook.md)|
|[Session](tasksmodule-session-property-outlook.md)|
|[Visible](tasksmodule-visible-property-outlook.md)|

## See also


#### Other resources


[Outlook Object Model Reference](http://msdn.microsoft.com/library/73221b13-d8d8-99b8-3394-b95dbbfd5ddc%28Office.15%29.aspx)
