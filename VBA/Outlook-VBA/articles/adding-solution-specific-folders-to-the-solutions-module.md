---
title: Adding Solution-Specific Folders to the Solutions Module
ms.prod: outlook
ms.assetid: 2180c3e3-b83b-7977-1bf6-61ae7cc64905
ms.date: 06/08/2017
---


# Adding Solution-Specific Folders to the Solutions Module

The  **Solutions** module is a navigation module that gives Microsoft Outlook solutions a way to expose one or more folders in the Navigation Pane. By default, its display name is **Solutions**, and its default position is '9' in the Navigation Pane, following the  **Tasks** navigation module.

To add a solution and its folders programmatically to the  **Solutions** module, you must use the **[AddSolution](solutionsmodule-addsolution-method-outlook.md)** method of the **[SolutionsModule](solutionsmodule-object-outlook.md)** object. You can customize the name of the module if there is only one solution in the module; otherwise the name reverts back to **Solutions**.

The  **Solutions** module shares some features with other navigation modules; for example, besides the module button itself, there is a smaller button for the module that users can click to collapse module buttons on the Navigation Pane. In addition, users can open the **Solutions** module in a new window and customize the display of modules and their relative positions on the Navigation Pane by using the shortcut menu.

The  **Solutions** module displays folders from each solution in its own group. Each solution has a corresponding solution root folder, and any subfolders under that root are displayed in ascending alphabetical order. Subfolders contain items of different item types; for example, you might have one subfolder for calendar items and another subfolder for tasks items. You can specify whether these folders appear under the default folders for their respective item types as well. Solution root folders are displayed in the chronological order that you added them to the **Solutions** module and cannot be reordered.
A solution root folder and its subfolders must reside on the same store. If the solution root folder is the root folder of a store, all subfolders of the store root folder are displayed under the solution root folder. Subfolders of the store root folder include the **Deleted Items** folder and the **Searches** root folder. The active/inactive state of search subfolders under the **Searches** root folder is the same in the **Solutions** module as it is in the **Folder List**.

## See also


#### Concepts


 [Customizing the Navigation Pane](customizing-the-navigation-pane.md)

