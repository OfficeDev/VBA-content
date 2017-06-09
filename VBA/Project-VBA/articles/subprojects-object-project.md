---
title: Subprojects Object (Project)
ms.prod: project-server
ms.assetid: 15688529-6d9c-6429-0d22-a5a16c033dcc
ms.date: 06/08/2017
---


# Subprojects Object (Project)

Contains a collection of  **[Subproject](subproject-object-project.md)** objects


## Example

 **Using the Subprojects Collection Object**

Use  **Subprojects** ( _Index_ ), where _Index_ is the subproject index or project summary task name, to return a single **Subproject** object. The following example prevents changes made to the specified subproject in a master project from being automatically made to the source project.




```
ActiveProject.Subprojects("Arcadia Bay Online Catalog Plan").LinkToSource = False
```

 **Getting the Subprojects Collection object**

Use the  **[Subprojects](http://msdn.microsoft.com/library/e4b143fb-3da7-69bd-6535-5604c2cc2dc0%28Office.15%29.aspx)** property to return a **Subprojects** collection. The following example cautions the user if any of the subprojects in the active project are not on the hard disk.




```
Dim SubProj As Subproject 

 

For Each SubProj in ActiveProject.Subprojects 

 If UCase(Left$(SubProj.Path, 1)) <> "C" Then 

 MsgBox Right$(SubProj.Path, InStrRev(SubProj.Path, "\") - 1) &amp; _ 

 " is not on your local hard disk.", vbExclamation 

 End If 

Next SubProj
```


## Properties



|**Name**|
|:-----|
|[Application](http://msdn.microsoft.com/library/ba8620ec-6380-94f0-a47d-faba9ba04fb4%28Office.15%29.aspx)|
|[Count](http://msdn.microsoft.com/library/ddbbcd5b-3885-fae9-14ef-4854d9d3874f%28Office.15%29.aspx)|
|[Item](http://msdn.microsoft.com/library/5044cc36-2e53-d424-c037-dbebe30d821a%28Office.15%29.aspx)|
|[Parent](http://msdn.microsoft.com/library/86af8044-cc92-fbf3-d98c-1d3b6ba7ca2a%28Office.15%29.aspx)|

## See also


#### Other resources


[Project Object Model](http://msdn.microsoft.com/library/900b167b-88ec-ea88-15b7-27bb90c22ac6%28Office.15%29.aspx)
