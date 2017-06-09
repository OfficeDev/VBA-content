---
title: Enumerate, Show, Hide, and Position Modules in the Navigation Pane
ms.prod: outlook
ms.assetid: 3e510798-3a31-6ec6-6c45-8e0d1759ca1b
ms.date: 06/08/2017
---


# Enumerate, Show, Hide, and Position Modules in the Navigation Pane

The  **[NavigationModules](navigationmodules-object-outlook.md)** property of the **[NavigationPane](navigationpane-object-outlook.md)** object in Microsoft Outlook provides access to the navigation modules contained by the Navigation Pane. You can use the **Item** method to enumerate the **[NavigationModule](navigationmodule-object-outlook.md)** objects contained by the collection, as the **Item** method is both the default property and the indexer property for the **NavigationModules** collection. The **[CurrentModule](navigationpane-currentmodule-property-outlook.md)** property determines which **NavigationModule** object is currently selected in the Navigation Pane.

Also, each  **NavigationModule** object provides several properties that can be used to show, hide, or change the display position of modules in the Navigation Pane:

- The  **[Visible](navigationmodule-visible-property-outlook.md)** property determines whether a **NavigationModule** object can be displayed in the Navigation Pane.
    
- The  **[Position](navigationmodule-position-property-outlook.md)** property determines the ordinal position of a **NavigationModule** object when displayed in the Navigation Pane.
    
The  **[DisplayedModuleCount](navigationpane-displayedmodulecount-property-outlook.md)** property of the **NavigationPane** object determines the number of visible **NavigationModule** objects that can be displayed by the Navigation Pane. If the **Visible** property of a **NavigationModule** object is set to **False**, or if the  **Position** property of the **NavigationModule** object is set such that the module doesn't fall within the number of visible **NavigationModule** objects that can be displayed in the Navigation Pane, the module isn't displayed.
The following code samples in Microsoft Visual Basic for Applications (VBA) consist of the  `MoveCurrentModuleToTop` and `MakeAllModulesVisible` procedures.
The  `MoveCurrentModuleToTop` procedure uses the **CurrentModule** property of the **NavigationPane** object to retrieve the currently selected **NavigationModule** object and sets the **Position** property of that **NavigationModule** object to 1, making it the topmost displayed module in the Navigation Pane.
The  `MoveCurrentModuleToTop` procedure enumerates the **Modules** collection of the **NavigationPane** object, setting the **Visible** property of each **NavigationModule** object contained in the collection to **True**. It finally sets the  **[DisplayedModuleCount](navigationpane-displayedmodulecount-property-outlook.md)** property of the **NavigationPane** object to the value of the **[Count](navigationmodules-count-property-outlook.md)** property of the **NavigationModules** collection for the **NavigationPane** object, ensuring that every navigation module contained in the Navigation Pane is visible to the user.



```vb
Private Sub MoveCurrentModuleToTop() 
 
 Dim objPane As NavigationPane 
 
 ' Get the NavigationPane object for the 
 ' currently displayed Explorer object. 
 Set objPane = Application.ActiveExplorer.NavigationPane 
 
 ' Set the Position property of the currently selected 
 ' module to 1, making it the topmost module displayed 
 ' in the Navigation Pane. 
 objPane.CurrentModule.Position = 1 
End Sub 
 
Private Sub MakeAllModulesVisible() 
 
 Dim objPane As NavigationPane 
 Dim objModule As NavigationModule 
 
 ' Get the NavigationPane object for the 
 ' currently displayed Explorer object. 
 Set objPane = Application.ActiveExplorer.NavigationPane 
 
 ' This loop enumerates through the Modules collection, 
 ' setting the Visible property of each module to True. 
 For Each objModule In objPane.Modules 
 objModule.Visible = True 
 Next 
 
 ' Set the DisplayedModuleCount property to 
 ' display all modules contained by the 
 ' Navigation Pane. 
 objPane.DisplayedModuleCount = objPane.Modules.count 
End Sub
```


