---
title: Groups Object (Project)
ms.prod: project-server
ms.assetid: 2e4c4846-6193-fc12-ad02-0dd69f88b31e
ms.date: 06/08/2017
---


# Groups Object (Project)

Represents a collection of  **[Group](group-object-project.md)** objects.
 


## Remarks

For groups where the group hierarchy can be maintained and cell color can be a hexadecimal value, use the  **[Groups2](groups2-object-project.md)** collection object.
 

 
Use  `TaskGroups(Index)` or ` ResourceGroups(Index)`, where *Index* is the group definition index or group definition name, to return a **Group** object.
 

 

## Example

The following example ensures that the Standard Rate resource group displays summary task information.
 

 

```
ActiveProject.ResourceGroups("Standard Rate").ShowSummary = True 


```


## Methods



|**Name**|
|:-----|
|[Add](groups-add-method-project.md)|
|[Copy](groups-copy-method-project.md)|

## Properties



|**Name**|
|:-----|
|[Application](groups-application-property-project.md)|
|[Count](groups-count-property-project.md)|
|[Item](groups-item-property-project.md)|
|[Parent](groups-parent-property-project.md)|

## See also


#### Other resources


 
[Project Object Model](http://msdn.microsoft.com/library/900b167b-88ec-ea88-15b7-27bb90c22ac6%28Office.15%29.aspx)
