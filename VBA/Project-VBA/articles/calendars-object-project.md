---
title: Calendars Object (Project)
ms.prod: project-server
ms.assetid: a96c7b96-f0ab-5ec3-3d16-facea61b8ee5
ms.date: 06/08/2017
---


# Calendars Object (Project)

Contains a collection of  **[Calendar](calendar-object-project.md)** objects.


## Example

 **Using the Calendar Object**

Use  **BaseCalendars(** _Index_ **)**, where _Index_ is the calendar index number or calendar name, to return a single **Calendar** object.




```
MsgBox ActiveProject.BaseCalendars(1).Name
```

 **Using the Calendars Collection**

Use the  **[BaseCalendars](http://msdn.microsoft.com/library/fb7f55f6-6618-fb82-dae1-320953bcf79d%28Office.15%29.aspx)** property to return a **Calendars** collection. The following example resets the properties of each base calendar in the active project to their default values.




```
Dim C As Calendar 

 

For Each C In ActiveProject.BaseCalendars 

 C.Reset 

Next C
```

Use the  **[BaseCalendarCreate](http://msdn.microsoft.com/library/c9c92dff-255a-041b-c18d-49d6d75884e3%28Office.15%29.aspx)** method to add a **Calendar** object to the **Calendars** collection. The following example creates a new base calendar.




```
BaseCalendarCreate Name:="Base Holiday Calendar"
```


## Properties



|**Name**|
|:-----|
|[Application](http://msdn.microsoft.com/library/8101846d-3996-8c44-12ad-ad63fc4ce094%28Office.15%29.aspx)|
|[Count](http://msdn.microsoft.com/library/a7652285-5694-4439-5cd9-ff691d29a6a2%28Office.15%29.aspx)|
|[Item](http://msdn.microsoft.com/library/de9595de-a159-e19a-6a7c-81c67ca7557f%28Office.15%29.aspx)|
|[Parent](http://msdn.microsoft.com/library/51bada64-c5db-c3af-5bf0-0979aec8bbc4%28Office.15%29.aspx)|

## See also


#### Other resources


[Project Object Model](http://msdn.microsoft.com/library/900b167b-88ec-ea88-15b7-27bb90c22ac6%28Office.15%29.aspx)
