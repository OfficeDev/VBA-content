---
title: Show or Hide the Navigation Pane
ms.prod: outlook
ms.assetid: ef4ad7b9-6475-7b28-ce79-fbefe29b193c
ms.date: 06/08/2017
---


# Show or Hide the Navigation Pane

You can set the  **[IsCollapsed](navigationpane-iscollapsed-property-outlook.md)** property of the **[NavigationPane](navigationpane-object-outlook.md)** object to collapse or expand the Navigation Pane for an **[Explorer](explorer-object-outlook.md)** object. The appearance of the Navigation Pane changes depending on the setting of the **IsCollapsed** property, as well as the settings of other properties for the **NavigationPane** object. The **[DisplayedModuleCount](navigationpane-displayedmodulecount-property-outlook.md)** property of the **NavigationPane** object determines the number of modules that can be displayed by the Navigation Pane, while the **[Visible](navigationmodule-visible-property-outlook.md)** and **[Position](navigationmodule-position-property-outlook.md)** property values of each **[NavigationModule](navigationmodule-object-outlook.md)** object determine which modules are displayed, and in what order.

Setting the  **IsCollapsed** property to **True** collapses the Navigation Pane. When collapsed, the Navigation Pane displays only the icon for each visible navigation module.

Setting the  **IsCollapsed** property to **False** expands the Navigation Pane. When expanded, the Navigation Pane displays the icon and name for the number of topmost visible modules contained in the **[NavigationModules](navigationmodules-object-outlook.md)** collection of the **NavigationPane** object, specified by the **DisplayedModuleCount** property. All other visible modules are displayed as icons at the bottom of the Navigation Pane.

For example, the  **NavigationModules** collection of an expanded **NavigationPane** object for the active explorer contains eight modules. All modules have a **Visible** property value of **True**, with the exception of the third navigation module (the  **Module** object with the **Position** property value set to 3.) If the **DisplayedModuleCount** property is set to 4, the icons and names of only the first four visible **NavigationModule** objects, with **Position** property values of 1, 2, 4, and 5, are displayed as large buttons in the Navigation Pane. The remaining three visible **NavigationModule** objects, in positions 6, 7, and 8, are displayed only as icons, on small buttons at the bottom of the Navigation Pane. If the **IsCollapsed** property is set to **False**, the collapsed Navigation Pane displays the first four visible  **NavigationModule** objects only as icons. The remaining three visible **NavigationModule** objects are available on the Navigation Pane dropdown menu.
The following sample ensures that the Navigation Pane is always expanded whenever the currently selected navigation module changes, either programmatically or by user action, by setting the  **IsCollapsed** property to **False**. The sample performs the following actions:

1. The sample first obtains a reference to the  **NavigationPane** object for the active explorer when the **[Startup](application-startup-event-outlook.md)** event of the **[Application](application-object-outlook.md)** object is raised and assigns it to `objPane`, so the  **[ModuleSwitch](navigationpane-moduleswitch-event-outlook.md)** event of the **NavigationPane** object can be detected.
    
2. When the  **ModuleSwitch** event of the **NavigationPane** occurs, the sample then checks if the current navigation module has changed by comparing the contents of the _CurrentModule_ parameter of the **ModuleSwitch** event against the **[CurrentModule](navigationpane-currentmodule-property-outlook.md)** property of the **NavigationPane** object. If these object references are different, the **IsCollapsed** property of the **NavigationPane** object is set to **False**.
    



```vb
Dim WithEvents objPane As NavigationPane 
 
Private Sub Application_Startup() 
 ' Get the NavigationPane object for the 
 ' currently displayed Explorer object. 
 Set objPane = Application.ActiveExplorer.NavigationPane 
End Sub 
 
Private Sub objPane_ModuleSwitch(ByVal CurrentModule As NavigationModule) 
 
 ' Check if the currently selected navigation module 
 ' has changed. 
 If Not (CurrentModule Is objPane.CurrentModule) Then 
 
 ' Set the IsCollapsed property to 
 ' ensure that the Navigation Pane 
 ' is visible. 
 If Not (objPane Is Nothing) Then 
 objPane.IsCollapsed = False 
 End If 
 End If 
 
End Sub
```


