---
title: Groups2 Object (Project)
ms.prod: project-server
ms.assetid: b2b83868-3366-4fb0-fed9-16d4c5eaff87
ms.date: 06/08/2017
---


# Groups2 Object (Project)

Represents a collection of  **[Group2](group2-object-project.md)** objects, which can maintain group hierarchy.
 


## Remarks

Use  `TaskGroups2(Index)` or `ResourceGroups2(Index)`, where *Index* is the group definition index or group definition name, to return a **Group2** object.
 

 

## Example

The following example ensures that the Standard Rate resource group displays summary task information.
 

 

```
ActiveProject.ResourceGroups2("Standard Rate").ShowSummary = True 


```


## Methods



|**Name**|
|:-----|
|[Add](groups2-add-method-project.md)|
|[Copy](groups2-copy-method-project.md)|

## Properties



|**Name**|
|:-----|
|[Application](groups2-application-property-project.md)|
|[Count](groups2-count-property-project.md)|
|[Item](groups2-item-property-project.md)|
|[Parent](groups2-parent-property-project.md)|

## See also


#### Other resources


 
[Project Object Model](http://msdn.microsoft.com/library/900b167b-88ec-ea88-15b7-27bb90c22ac6%28Office.15%29.aspx)
