
# Set a Module as the Currently Selected Module in the Navigation Pane

 **Last modified:** July 28, 2015

 _**Applies to:** Outlook 2013_

You can use the  ** [CurrentModule](df7086b3-4174-839f-0756-a5201379ed92.md)** property of the ** [NavigationPane](b6538c72-6115-99fc-c926-e0532a747823.md)** object in Microsoft Outlook to set a ** [NavigationModule](76565eaf-1e64-f5d4-b90f-ba156863802c.md)** object as the currently selected navigation module in the Navigation Pane of an ** [Explorer](026591e5-049f-503a-4166-34e6dbc225fb.md)** object.

The following sample sets the  **Calendar** navigation module as the currently selected navigation module if the **Journal** navigation module is selected, either programmatically or by user action, in the Navigation Pane. The sample performs the following actions:

1. The sample first obtains a reference to the  **NavigationPane** object for the active explorer when the ** [Startup](d4724d96-2572-b1e3-e202-0bfffb5cf7d5.md)** event of the ** [Application](797003e7-ecd1-eccb-eaaf-32d6ddde8348.md)** object is raised and assigns it to `objPane`, so the  ** [ModuleSwitch](63ecb01e-56e2-cfa8-0481-b81761f6ab5c.md)** event of the **NavigationPane** object can be detected.
    
2. When the  **ModuleSwitch** event of the **NavigationPane** occurs, the sample then checks if the current navigation module has changed by comparing the contents of the _CurrentModule_ parameter of the **ModuleSwitch** event against the ** [CurrentModule](df7086b3-4174-839f-0756-a5201379ed92.md)** property of the **NavigationPane** object.
    
3. If these object references are different, the sample then checks the  ** [NavigationModuleType](ee1fc78a-9720-c8d0-964c-0178ddbe8af6.md)** property of the **NavigationModule** object reference in the _CurrentModule_ parameter of the **ModuleSwitch** event.
    
4. If the  **NavigationModuleType** property of the currently selected **Module** object is set to **olModuleJournal**, the sample then displays a dialog box to indicate to the user that the currently selected  **Journal** navigation module is temporarily unavailable, and that instead the **Calendar** navigation module will be selected.
    
5. Finally, the sample uses the  ** [GetNavigationModule](7c1a1313-94a4-fa68-7e70-66d85496fec0.md)** method of the **Modules** collection for the **NavigationPane** object to attempt to retrieve a ** [CalendarModule](9203024d-9cef-75e0-600f-f3899e24761a.md)** object. If successful, the **CurrentModule** property of the **NavigationPane** object is set to the retrieved **CalendarModule** object reference.
    



```
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
 MsgBox "The Journal module is temporarily unavailable. " &amp; _ 
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

