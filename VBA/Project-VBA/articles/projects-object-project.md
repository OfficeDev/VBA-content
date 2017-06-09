---
title: Projects Object (Project)
keywords: vbapj.chm131311
f1_keywords:
- vbapj.chm131311
ms.prod: project-server
ms.assetid: 5a254428-f50d-e74f-dd31-5cdb260a4364
ms.date: 06/08/2017
---


# Projects Object (Project)

Contains a collection of **[Project](project-object-project.md)** objects.


## Example

 **Using the Project Object**

Use  **Projects** (Index), where Index is the project index number or project name, to return a single **Project** object. The following example switches among all the open projects, memorizes the full name of each, and then displays the results.




```
Dim Temp As Long, Names As String 

 

For Temp = 1 To Projects.Count 

 Projects(Temp).Activate 

 Names = Names &amp; Projects(Temp).FullName &amp; vbCrLf 

Next Temp 

 

MsgBox Names
```

 **Using the Projects Collection**

Use the  **[Projects](http://msdn.microsoft.com/library/792b7334-a424-abe1-287e-285d3ab362c7%28Office.15%29.aspx)** property to return a **Projects** collection. The following example counts the number of open projects.




```
Application.Projects.Count
```

Because the  **Projects** collection is a top-level object, the following example is functionally identical to the preceding one.




```
Projects.Count
```

Use the  **[Add](http://msdn.microsoft.com/library/51629c33-1521-bfee-edf7-bed792d393c1%28Office.15%29.aspx)** method to add a **Project** object to the **Projects** collection. The following example creates a new project without prompting for project information.




```
Projects.Add False
```


## Methods



|**Name**|
|:-----|
|[Add](http://msdn.microsoft.com/library/51629c33-1521-bfee-edf7-bed792d393c1%28Office.15%29.aspx)|
|[CanCheckOut](http://msdn.microsoft.com/library/330f28a3-d785-ae5d-0f64-8e02ac52d8d6%28Office.15%29.aspx)|
|[CheckOut](http://msdn.microsoft.com/library/2de8fef7-150b-4f67-4677-507f5d2a258f%28Office.15%29.aspx)|

## Properties



|**Name**|
|:-----|
|[Application](http://msdn.microsoft.com/library/cbba5bfd-63d5-97da-1fca-8ea4ca8ac7cf%28Office.15%29.aspx)|
|[Count](http://msdn.microsoft.com/library/e6b9ee18-36f1-4626-569b-ef03804e86b4%28Office.15%29.aspx)|
|[Item](http://msdn.microsoft.com/library/ec05fd24-c6b3-d3b8-d81c-1c4e0ad1d8ce%28Office.15%29.aspx)|
|[Parent](http://msdn.microsoft.com/library/0d892acb-413a-0765-1257-3bad4d3c7b67%28Office.15%29.aspx)|

## See also


#### Other resources


[Project Object Model](http://msdn.microsoft.com/library/900b167b-88ec-ea88-15b7-27bb90c22ac6%28Office.15%29.aspx)
