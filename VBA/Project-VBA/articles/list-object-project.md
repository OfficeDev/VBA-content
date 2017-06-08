---
title: List Object (Project)
ms.prod: project-server
api_name:
- Project.List
ms.assetid: 3934c2e8-d810-6571-9a33-1d41edbab87a
ms.date: 06/08/2017
---


# List Object (Project)

Represents a collection of strings or numbers that contain field identification numbers, field names, reports, resource filters, resource tables, resource views, task filters, task tables, task views, or views. (There is no collection for  **List** objects.) It can be accessed through the **List** properties of the appropriate objects.


## Example

 **Using the List Object**

Use a property such as the  **[ReportList](http://msdn.microsoft.com/library/0c688797-21cc-eaa0-0ebf-95e1e053f222%28Office.15%29.aspx)** property to return a **List** object. The following example displays a list of all the reports available in the active project.




```
Dim Items As Integer, ReportNames As String 
 
For Items = 1 To ActiveProject.ReportList.Count 
 ReportNames = ActiveProject.ReportList(Items) &amp; _ 
 ListSeparator &amp; " " &amp; ReportNames 
Next Items 
 
MsgBox Left$(ReportNames, Len(ReportNames) - Len(ListSeparator &amp; " "))
```


## Properties



|**Name**|
|:-----|
|[Application](http://msdn.microsoft.com/library/55f48bb5-e5cc-8117-9e01-be55964690af%28Office.15%29.aspx)|
|[Count](http://msdn.microsoft.com/library/d417857d-99f9-3c82-f211-4dd0241deb44%28Office.15%29.aspx)|
|[Item](http://msdn.microsoft.com/library/9dbe7805-82b7-650a-28c4-ec4d22914f66%28Office.15%29.aspx)|
|[Parent](http://msdn.microsoft.com/library/08d2d7d8-fafc-8f60-be78-c2d462005eaf%28Office.15%29.aspx)|

