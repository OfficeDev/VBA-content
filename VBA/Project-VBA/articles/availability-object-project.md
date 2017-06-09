---
title: Availability Object (Project)
ms.prod: project-server
api_name:
- Project.Availability
ms.assetid: 2b832aed-2b58-f020-2a2c-8756ec7ec1a4
ms.date: 06/08/2017
---


# Availability Object (Project)


 

Represents a line from the  **Resource Availability** grid for a resource. The **Availability** object is a member of the **[Availabilities](availabilities-object-project.md)** collection.
 
 **Using the Availability Object**
 
Use  **Availabilities(***Index* **)**, where*Index* is the availability index number, to return a single **Availability** object. The following example returns the availability information from the first line of the **Resource Availability** grid for the specified resource.
 



```
MsgBox ActiveProject.Resources("Tom").Name &amp; " is available from " &amp; _ 
    ActiveProject.Resources("Tom").Availabilities(1).AvailableFrom &amp; " to " &amp; _ 
    ActiveProject.Resources("Tom").Availabilities(1).AvailableTo &amp; "." 

```

Use the  **[Availabilities](resource-availabilities-property-project.md)** property to return an **Availabilities** collection. The following example displays the range of dates during which the specified resource is available for work.
 



```
Dim Avail As Availability 
 
For Each Avail In ActiveProject.Resources("Tom").Availabilities 
    MsgBox "From " &amp; Avail.AvailableFrom &amp; " to " &amp; Avail.AvailableTo 
Next Avail 

```

Use the  **[Add](availabilities-add-method-project.md)** method to add an **Availability** object to the **Availabilities** collection. The following example adds a line to the **Resource Availability** grid showing that the specified resource is available only half-time during the month of April.
 



```
ActiveProject.Resources("Tom").Availabilities.Add "4/1/2012", "4/30/2012", 50
```


## Methods



|**Name**|
|:-----|
|[Delete](availability-delete-method-project.md)|

## Properties



|**Name**|
|:-----|
|[Application](availability-application-property-project.md)|
|[AvailableFrom](availability-availablefrom-property-project.md)|
|[AvailableTo](availability-availableto-property-project.md)|
|[AvailableUnit](availability-availableunit-property-project.md)|
|[Index](availability-index-property-project.md)|
|[Parent](availability-parent-property-project.md)|

