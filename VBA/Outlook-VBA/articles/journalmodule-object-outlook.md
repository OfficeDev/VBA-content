---
title: JournalModule Object (Outlook)
keywords: vbaol11.chm3197
f1_keywords:
- vbaol11.chm3197
ms.prod: outlook
api_name:
- Outlook.JournalModule
ms.assetid: 5a696d10-8a10-c01d-cf65-f8a65718f120
ms.date: 06/08/2017
---


# JournalModule Object (Outlook)

Represents the  **Journal** navigation module in the Navigation Pane of an explorer.


## Remarks

The  **JournalModule** object, derived from the **[NavigationModule](navigationmodule-object-outlook.md)** object, provides access to the navigation groups contained in the **Journal** navigation module of the Navigation Pane for an explorer. Use the **[GetNavigationModule](navigationmodules-getnavigationmodule-method-outlook.md)** method or the **[Item](navigationmodules-item-method-outlook.md)** method of the **[Modules](navigationpane-modules-property-outlook.md)** collection for the parent **[NavigationPane](navigationpane-object-outlook.md)** object to retrieve a **NavigationModule** object, then use the **[NavigationModuleType](navigationmodule-navigationmoduletype-property-outlook.md)** property of the **NavigationModule** object to retrieve the module type. If the **NavigationModuleType** property is set to **olModuleJournal**, you can then cast the **NavigationModule** object reference as a **JournalModule** object to access the **[NavigationGroups](journalmodule-navigationgroups-property-outlook.md)** property for that navigation module.

You can use the  **[Visible](journalmodule-visible-property-outlook.md)** property to determine if the navigation module is visible and the **[Position](journalmodule-position-property-outlook.md)** property to return or set the display position of the navigation module within the Navigation Pane. You can use the **[Name](journalmodule-name-property-outlook.md)** property to return the display name of the **Journal** navigation module within the Navigation Pane.


## Properties



|**Name**|
|:-----|
|[Application](journalmodule-application-property-outlook.md)|
|[Class](journalmodule-class-property-outlook.md)|
|[Name](journalmodule-name-property-outlook.md)|
|[NavigationGroups](journalmodule-navigationgroups-property-outlook.md)|
|[NavigationModuleType](journalmodule-navigationmoduletype-property-outlook.md)|
|[Parent](journalmodule-parent-property-outlook.md)|
|[Position](journalmodule-position-property-outlook.md)|
|[Session](journalmodule-session-property-outlook.md)|
|[Visible](journalmodule-visible-property-outlook.md)|

## See also


#### Other resources


[Outlook Object Model Reference](http://msdn.microsoft.com/library/73221b13-d8d8-99b8-3394-b95dbbfd5ddc%28Office.15%29.aspx)
