---
title: Resources Object (Project)
ms.prod: project-server
ms.assetid: 84f8357a-358b-f2ae-e164-65c0c5abd383
ms.date: 06/08/2017
---


# Resources Object (Project)

Contains a collection of  **[Resource](resource-object-project.md)** objects.


## Example

 **Using the Resources Collection**

Use  **Resources** ( _Index_ ), where _Index_ is the resource index number or resource name, to return a single **Resource** object. The following example lists the names of all resources in the active project.




```
Dim R As Long, Names As String 

 

For R = 1 To ActiveProject.Resources.Count 

 Names = ActiveProject.Resources(R).Name &amp; ", " &amp; Names 

Next R 

 

Names = Left$(Names, Len(Names) - Len(ListSeparator &amp; " ")) 

MsgBox Names
```

 **Using the Resources Collection**

Use the  **[Resources](http://msdn.microsoft.com/library/40744aba-2b61-2b45-133a-f1dd9c7d6add%28Office.15%29.aspx)** property to return a **Resources** collection. The following example generates the same list as the previous example, but does so by setting an object reference to `ActiveProject.Resources` , and then using `R` where `ActiveProject.Resources` is used.




```
Dim R As Resources, Temp As Long, Names As String 

 

Set R = ActiveProject.Resources 

 

For Temp = 1 To R.Count 

 Names = R(Temp).Name &amp; ", " &amp; Names 

Next Temp 

 

Names = Left$(Names, Len(Names) - Len(ListSeparator &amp; " ")) 

MsgBox Names
```

Use the  **[Add](http://msdn.microsoft.com/library/4fb69f50-4ba6-89a4-f586-3df268ae7fd5%28Office.15%29.aspx)** method to add a **Resource** object to the **Resources** collection. The following example adds a new resource named Matilda to the active project.




```
ActiveProject.Resources.Add "Matilda"
```


## Methods



|**Name**|
|:-----|
|[Add](http://msdn.microsoft.com/library/4fb69f50-4ba6-89a4-f586-3df268ae7fd5%28Office.15%29.aspx)|

## Properties



|**Name**|
|:-----|
|[Application](http://msdn.microsoft.com/library/dbfa8ee9-4bae-c058-d940-eea2018f463d%28Office.15%29.aspx)|
|[Count](http://msdn.microsoft.com/library/53a50d7d-beea-2bed-c2dd-67b402a27e0c%28Office.15%29.aspx)|
|[Item](http://msdn.microsoft.com/library/2c0c95b0-07fa-a8b8-05a3-50072824c8a8%28Office.15%29.aspx)|
|[Parent](http://msdn.microsoft.com/library/138d0de6-c374-6f7d-0e4d-6bb515ce8c4e%28Office.15%29.aspx)|
|[UniqueID](http://msdn.microsoft.com/library/84c48d8e-45e7-f1d7-9284-cb7f92c3ffb0%28Office.15%29.aspx)|

## See also


#### Other resources


[Project Object Model](http://msdn.microsoft.com/library/900b167b-88ec-ea88-15b7-27bb90c22ac6%28Office.15%29.aspx)
