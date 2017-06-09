---
title: CalendarModule.Position Property (Outlook)
keywords: vbaol11.chm2828
f1_keywords:
- vbaol11.chm2828
ms.prod: outlook
api_name:
- Outlook.CalendarModule.Position
ms.assetid: 3857d981-acd7-975c-0ff1-453ee2b7402e
ms.date: 06/08/2017
---


# CalendarModule.Position Property (Outlook)

Returns or sets a  **Long** value that represents the ordinal position of the **[CalendarModule](calendarmodule-object-outlook.md)** object when it is displayed in the Navigation Pane. Read/write.


## Syntax

 _expression_ . **Position**

 _expression_ A variable that represents a **CalendarModule** object.


## Remarks

This property can only be set to a value between 1 and 9. An error occurs if you attempt to set it to a value outside of that range.

Changing the value of this property for a given  **CalendarModule** object changes the **Position** values of other navigation modules in a **[NavigationModules](navigationmodules-object-outlook.md)** collection, depending on the relative change between the new value and the original value.


- If the new value is less than the original value, the specified  **CalendarModule** object moves up to the new position and the other navigation modules that are already at or below that new position move down.
    
- If the new value is greater than the original value, the specified  **CalendarModule** object moves down to the new position and the other navigation modules that are between the old position and the new position move up, filling the old position.
    

## Example

The following Visual Basic for Applications (VBA) sample code attempts to retrieve the  **Calendar** navigation module from the Navigation Pane. If it successfully retrieves the module, the code sets the **Position** property of the **CalendarModule** object to '1,' which moves it to the top of the Navigation Pane. Finally, the code sets the **[CurrentModule](navigationpane-currentmodule-property-outlook.md)** property of the **[NavigationPane](navigationpane-object-outlook.md)** object to the retrieved **Calendar** module, which selects it in the Navigation Pane.


```vb
Sub MoveCalendarModuleFirst() 
 
 Dim objPane As NavigationPane 
 
 Dim objModule As CalendarModule 
 
 
 
 On Error GoTo ErrRoutine 
 
 
 
 ' Get the current NavigationPane object. 
 
 Set objPane = Application.ActiveExplorer.NavigationPane 
 
 
 
 ' Get the Calendar navigation module 
 
 ' from the Navigation Pane. 
 
 Set objModule = objPane.Modules.GetNavigationModule( _ 
 
 olModuleCalendar) 
 
 
 
 ' If a CalendarModule object is present, 
 
 ' make it the first navigation module displayed in the 
 
 ' Navigation Pane. 
 
 If Not (objModule Is Nothing) Then 
 
 objModule.Position = 1 
 
 End If 
 
 
 
 ' Select the Calendar navigation module in the 
 
 ' Navigation Pane. 
 
 Set objPane.CurrentModule = objModule 
 
 
 
EndRoutine: 
 
 On Error GoTo 0 
 
 Set objModule = Nothing 
 
 Set objPane = Nothing 
 
 Exit Sub 
 
 
 
ErrRoutine: 
 
 Debug.Print Err.Number &; " (&;H" &; Hex(Err.Number) &; ")" 
 
 Select Case Err.Number 
 
 Case -2147024809 '&;H80070057 
 
 ' Typically occurs if you set the Position 
 
 ' property less than 1 or greater than 9. 
 
 MsgBox Err.Number &; " - " &; Err.Description, _ 
 
 vbOKOnly Or vbCritical, _ 
 
 "MoveCalendarModuleFirst" 
 
 End Select 
 
 GoTo EndRoutine 
 
End Sub
```


## See also


#### Concepts


[CalendarModule Object](calendarmodule-object-outlook.md)

