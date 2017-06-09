---
title: ResourceGroups Object (Project)
ms.prod: project-server
ms.assetid: 37bd0f3a-4d0e-1311-4409-ed31e0fe2e3a
ms.date: 06/08/2017
---


# ResourceGroups Object (Project)


 

Represents all of the resource-based group definitions.  **ResourceGroups** is a collection of **[Group](group-object-project.md)** objects.
 
 **Using the ResourceGroups Collection**
 
Use the  **[ResourceGroups](project-resourcegroups-property-project.md)** property to return a **ResourceGroups** collection. The following example lists the names of all the resource groups in the active project.
 



```
Dim rg As Group 
Dim rGroups As String 
 
For Each rg in ActiveProject.ResourceGroups 
 rGroups = rGroups &amp; rg.Name &amp; vbCrLf 
Next rg 
 
MsgBox rGroups
```

Use the  **[Add](resourcegroups-add-method-project.md)** method to add a **Group** object to the **ResourceGroups** collection. The following example creates a new group that groups resources by their standard rate and then modifies the criterion so that the resources are sorted in descending order.
 



```
ActiveProject.ResourceGroups.Add "Resources by Rate", "Standard Rate" 
ActiveProject.ResourceGroups("Resources by Rate").GroupCriteria(1).Ascending = False
```


## Remarks

For resource groups where the group hierarchy can be maintained and cell color can be a hexadecimal value, use the  **[ResourceGroups2](resourcegroups2-object-project.md)** collection object.
 

 

## Methods



|**Name**|
|:-----|
|[Add](resourcegroups-add-method-project.md)|
|[Copy](resourcegroups-copy-method-project.md)|

## Properties



|**Name**|
|:-----|
|[Application](resourcegroups-application-property-project.md)|
|[Count](resourcegroups-count-property-project.md)|
|[Item](resourcegroups-item-property-project.md)|
|[Parent](resourcegroups-parent-property-project.md)|

