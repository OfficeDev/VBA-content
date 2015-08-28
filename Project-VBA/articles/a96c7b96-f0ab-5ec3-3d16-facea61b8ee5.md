
# Calendars Object (Project)

 **Last modified:** July 28, 2015

Contains a collection of  ** [Calendar](2d3b0f05-4762-0058-15d4-47e1d2b9d9a9.md)** objects.

## Example

 **Using the Calendar Object**

Use  **BaseCalendars(**_Index_**)** , where _Index_ is the calendar index number or calendar name, to return a single **Calendar** object.




```
MsgBox ActiveProject.BaseCalendars(1).Name
```

 **Using the Calendars Collection**

Use the  ** [BaseCalendars](fb7f55f6-6618-fb82-dae1-320953bcf79d.md)** property to return a **Calendars** collection. The following example resets the properties of each base calendar in the active project to their default values.




```
Dim C As Calendar 

 

For Each C In ActiveProject.BaseCalendars 

 C.Reset 

Next C
```

Use the  ** [BaseCalendarCreate](c9c92dff-255a-041b-c18d-49d6d75884e3.md)** method to add a **Calendar** object to the **Calendars** collection. The following example creates a new base calendar.




```
BaseCalendarCreate Name:="Base Holiday Calendar"
```


## See also


#### Concepts


 [Project Object Model](900b167b-88ec-ea88-15b7-27bb90c22ac6.md)
