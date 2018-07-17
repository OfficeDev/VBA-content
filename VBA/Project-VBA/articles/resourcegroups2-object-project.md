---
title: ResourceGroups2 Object (Project)
ms.prod: project-server
ms.assetid: b1328c39-42bc-4e9b-e268-1f308cd7ebb1
ms.date: 06/08/2017
---


# ResourceGroups2 Object (Project)

Represents all of the resource-based group definitions, where group hierarchy can be maintained.  **ResourceGroups2** is a collection of **[Group2](group2-object-project.md)** objects.
 


## Example

 **Using the ResourceGroups2 Collection**
 

 
Use the  **[ResourceGroups2](project-resourcegroups2-property-project.md)** property to return a **ResourceGroups2** collection. The following example lists the names of all the resource groups in the active project.
 

 



```
Dim rg2 As Group2  
Dim rGroups2 As String  
  
For Each rg2 in ActiveProject.ResourceGroups2  
    rGroups2 = rGroups2 &amp; rg2.Name &amp; vbCrLf  
Next rg2  
  
MsgBox rGroups2
```

Use the  **[Add](resourcegroups2-add-method-project.md)** method to add a **Group2** object to the **ResourceGroups2** collection. The following example creates a new group that groups resources by their standard rate and then modifies the criterion so that the resources are sorted in descending order.
 

 



```
ActiveProject.ResourceGroups2.Add "Resources by Rate", "Standard Rate"  
ActiveProject.ResourceGroups2("Resources by Rate").GroupCriteria(1).Ascending = False
```


## Methods



|**Name**|
|:-----|
|[Add](resourcegroups2-add-method-project.md)|
|[Copy](resourcegroups2-copy-method-project.md)|

## Properties



|**Name**|
|:-----|
|[Application](resourcegroups2-application-property-project.md)|
|[Count](resourcegroups2-count-property-project.md)|
|[Item](resourcegroups2-item-property-project.md)|
|[Parent](resourcegroups2-parent-property-project.md)|

## See also


#### Other resources


 
[Project Object Model](http://msdn.microsoft.com/library/900b167b-88ec-ea88-15b7-27bb90c22ac6%28Office.15%29.aspx)
