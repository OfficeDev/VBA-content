---
title: Display Specific Modules in the Navigation Pane
ms.prod: outlook
ms.assetid: 1a1017da-3047-fd58-fd92-ce0e750df7a6
ms.date: 06/08/2017
---


# Display Specific Modules in the Navigation Pane

The  **[CurrentModule](navigationpane-currentmodule-property-outlook.md)** property of the **[NavigationPane](navigationpane-object-outlook.md)** object, in Microsoft Outlook, determines which navigation module is currently displayed in the Navigation Pane. You can retrieve a reference to a given **[NavigationModule](navigationmodule-object-outlook.md)** object from a **NavigationPane** object by either enumerating the **[NavigationModules](navigationmodules-object-outlook.md)** collection or by using the **[GetNavigationModule](navigationmodules-getnavigationmodule-method-outlook.md)** method of the **NavigationModules** collection.

The following sample sets the  **Mail** navigation module as the currently selected navigation module if any navigation module is selected, either programmatically or by user action, in the Navigation Pane. The sample performs the following actions:

1. The sample first obtains a reference to the  **NavigationPane** object for the active explorer when the **[Startup](application-startup-event-outlook.md)** event of the **[Application](application-object-outlook.md)** object is raised and assigns it to `objPane`, so the  **[ModuleSwitch](navigationpane-moduleswitch-event-outlook.md)** event of the **NavigationPane** object can be detected.
    
2. When the  **ModuleSwitch** event of the **NavigationPane** occurs, the sample then checks the **[NavigationModuleType](navigationmodule-navigationmoduletype-property-outlook.md)** property of the **NavigationModule** object reference in the _CurrentModule_ parameter of the **ModuleSwitch** event.
    
3. If the  **NavigationModuleType** property of the currently selected **NavigationModule** object is set to **olModuleMail**, the sample uses the  **GetNavigationModule** method of the **NavigationModules** collection for the **NavigationPane** object to attempt to retrieve a **[MailModule](mailmodule-object-outlook.md)** object. If successful, the sample finally sets the **CurrentModule** property of the **NavigationPane** object to the retrieved **MailModule** object reference.
    



```vb
Dim WithEvents objPane As NavigationPane 
 
Private Sub Application_Startup() 
 ' Get the NavigationPane object for the 
 ' currently displayed Explorer object. 
 Set objPane = Application.ActiveExplorer.NavigationPane 
End Sub 
 
Private Sub objPane_ModuleSwitch(ByVal CurrentModule As NavigationModule) 
 Dim objModule As MailModule 
 
 If CurrentModule.NavigationModuleType <> olModuleMail Then 
 ' Use the GetModule method to obtain a MailModule from 
 ' the current NavigationPane object. 
 Set objModule = objPane.Modules.GetNavigationModule(olModuleMail) 
 
 ' Set the CurrentModule property to the MailModule. 
 Set objPane.CurrentModule = objModule 
 End If 
End Sub
```


