---
title: Set a Module as the Currently Selected Module in the Navigation Pane
ms.prod: outlook
ms.assetid: c7aeafcf-d88d-8d79-8dfd-e336cf00f101
ms.date: 06/08/2017
---


# Set a Module as the Currently Selected Module in the Navigation Pane

You can use the  **[CurrentModule](navigationpane-currentmodule-property-outlook.md)** property of the **[NavigationPane](navigationpane-object-outlook.md)** object in Microsoft Outlook to set a **[NavigationModule](navigationmodule-object-outlook.md)** object as the currently selected navigation module in the Navigation Pane of an **[Explorer](explorer-object-outlook.md)** object.

The following sample sets the  **Calendar** navigation module as the currently selected navigation module if the **Journal** navigation module is selected, either programmatically or by user action, in the Navigation Pane. The sample performs the following actions:

1. The sample first obtains a reference to the  **NavigationPane** object for the active explorer when the **[Startup](application-startup-event-outlook.md)** event of the **[Application](application-object-outlook.md)** object is raised and assigns it to `objPane`, so the  **[ModuleSwitch](navigationpane-moduleswitch-event-outlook.md)** event of the **NavigationPane** object can be detected.
    
2. When the  **ModuleSwitch** event of the **NavigationPane** occurs, the sample then checks if the current navigation module has changed by comparing the contents of the _CurrentModule_ parameter of the **ModuleSwitch** event against the **[CurrentModule](navigationpane-currentmodule-property-outlook.md)** property of the **NavigationPane** object.
    
3. If these object references are different, the sample then checks the  **[NavigationModuleType](navigationmodule-navigationmoduletype-property-outlook.md)** property of the **NavigationModule** object reference in the _CurrentModule_ parameter of the **ModuleSwitch** event.
    
4. If the  **NavigationModuleType** property of the currently selected **Module** object is set to **olModuleJournal**, the sample then displays a dialog box to indicate to the user that the currently selected  **Journal** navigation module is temporarily unavailable, and that instead the **Calendar** navigation module will be selected.
    
5. Finally, the sample uses the  **[GetNavigationModule](navigationmodules-getnavigationmodule-method-outlook.md)** method of the **Modules** collection for the **NavigationPane** object to attempt to retrieve a **[CalendarModule](calendarmodule-object-outlook.md)** object. If successful, the **CurrentModule** property of the **NavigationPane** object is set to the retrieved **CalendarModule** object reference.
    



```vb
Dim WithEvents objPane As NavigationPane 
 
Private Sub Application_Startup() 
 ' Get the NavigationPane object for the 
 ' currently displayed Explorer object. 
 Set objPane = Application.ActiveExplorer.NavigationPane 
 
End Sub 
 
Private Sub objPane_ModuleSwitch(ByVal CurrentModule As NavigationModule) 
 Dim objModule As CalendarModule 
 
 ' Check if the currently selected navigation module 
 ' has changed. 
 If Not (CurrentModule Is objPane.CurrentModule) Then 
 ' If the Journal module was selected, forcibly change 
 ' it to the Calendar module by setting the 
 ' CurrentModule property of the NavigationPane object. 
 If CurrentModule.NavigationModuleType = olModuleJournal Then 
 
 ' Let the user know what's happening. 
 MsgBox "The Journal module is temporarily unavailable. " &; _ 
 " Outlook is switching to the Calendar module, if available." 
 
 ' Retrieve the Calendar module, if one exists, for the 
 ' current Navigation Pane. 
 Set objModule = objPane.Modules.GetNavigationModule(olModuleCalendar) 
 
 ' If we have one, set the CurrentModule property of the 
 ' NavigationPane object to the Calendar module. 
 If Not (objModule Is Nothing) Then 
 Set objPane.CurrentModule = objModule 
 End If 
 End If 
 End If 
 
End Sub
```


