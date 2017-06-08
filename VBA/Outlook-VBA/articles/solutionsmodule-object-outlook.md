---
title: SolutionsModule Object (Outlook)
keywords: vbaol11.chm3371
f1_keywords:
- vbaol11.chm3371
ms.prod: outlook
api_name:
- Outlook.SolutionsModule
ms.assetid: 4597765e-a95d-bf07-2ac4-103218ebc696
ms.date: 06/08/2017
---


# SolutionsModule Object (Outlook)

Represents the  **Solutions** navigation module in the Navigation Pane of an explorer.


## Remarks

The  **Solutions** navigation module contains folders that developers of individual add-ins want to expose to users in the Navigation Pane. Each solution has one root folder under the **Solutions** module, and each root folder can contain subfolders that hold heterogeneous Outlook items.

To add solution folders programmatically to the  **Solutions** module, use the **SolutionsModule** object, which is derived from the **[NavigationModule](navigationmodule-object-outlook.md)** object.

To obtain an object for the  **Solutions** module, you must first determine whether the **Solutions** module exists in the Navigation Pane. To do that, use the **Modules** property for the **[NavigationPane](navigationpane-object-outlook.md)** object to obtain a **[NavigationModules](navigationmodules-object-outlook.md)** collection, and then specify the argument **olModuleSolutions** in the **[GetNavigationModule](navigationmodules-getnavigationmodule-method-outlook.md)** method of the **NavigationModules** collection.

If the call is successful, you can then cast the returned  **NavigationModule** object reference as a **SolutionsModule** object to access the properties and methods for that navigation module.

To add a solution root folder and its subfolders, pass a  **[Folder](folder-object-outlook.md)** object reference to the **[AddSolution](solutionsmodule-addsolution-method-outlook.md)** method of the **SolutionsModule** object. The default position of the **Solutions** module on the Navigation Pane is '9'.

If no solutions have been added to the  **Solutions** module, it is not visible in the Navigation Pane, and any attempt to set the **[Position](solutionsmodule-position-property-outlook.md)** or the **[Visible](solutionsmodule-visible-property-outlook.md)** properties of the **SolutionsModule** object raises an error. In addition, any attempt to set the **SolutionsModule** as the **[CurrentModule](navigationpane-currentmodule-property-outlook.md)** property of the **NavigationPane** object raises an error.


## Example

To see an example of an add-in that adds folders to the  **Solutions** module, see the article[Programming the Outlook 2010 Solutions Module](http://msdn.microsoft.com/en-us/library/ee692173%28office.14%29.aspx) on MSDN. The add-in in the article renames the **Solutions** module as **Solution Demo**, adds calendar, contacts, and tasks folders as subfolders to the solution root folder, sets custom icons for each of the subfolders, and customizes the Navigation Pane to move and enlarge the button for the  **Solution Demo** module.


## Methods



|**Name**|
|:-----|
|[AddSolution](solutionsmodule-addsolution-method-outlook.md)|

## Properties



|**Name**|
|:-----|
|[Application](solutionsmodule-application-property-outlook.md)|
|[Class](solutionsmodule-class-property-outlook.md)|
|[Name](solutionsmodule-name-property-outlook.md)|
|[NavigationModuleType](solutionsmodule-navigationmoduletype-property-outlook.md)|
|[Parent](solutionsmodule-parent-property-outlook.md)|
|[Position](solutionsmodule-position-property-outlook.md)|
|[Session](solutionsmodule-session-property-outlook.md)|
|[Visible](solutionsmodule-visible-property-outlook.md)|

## See also


#### Other resources


[Outlook Object Model Reference](http://msdn.microsoft.com/library/73221b13-d8d8-99b8-3394-b95dbbfd5ddc%28Office.15%29.aspx)
