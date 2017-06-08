---
title: GroupCriteria Object (Project)
ms.prod: project-server
ms.assetid: b19beefb-bfe2-54ba-0835-11624e92bafc
ms.date: 06/08/2017
---


# GroupCriteria Object (Project)

Contains a collection of  **[GroupCriterion](groupcriterion-object-project.md)** objects.
 


## Remarks

For groups where the group hierarchy can be maintained and cell color can be a hexadecimal value, use the  **[GroupCriteria2](groupcriteria2-object-project.md)** collection object.
 

 

## Example

 **Using the GroupCriterion Object**
 

 
Use  **GroupCriteria(***Index* **)**, where*Index* is the criterion index, to return a single **GroupCriterion** object. The following example sets the cell color for the first criterion in the Standard Rate resource group to blue.
 

 



```
ActiveProject.ResourceGroups("Standard Rate").GroupCriteria(1).CellColor = pjBlue
```

 **Using the GroupCriteria Collection**
 

 
Use the  **[GroupCriteria](group-groupcriteria-property-project.md)** property to return a **GroupCriteria** collection. The following example displays a list of the fields used as criteria in the specified task group and whether they are sorted in ascending or descending order.
 

 



```
Dim GC As GroupCriterion 

Dim Fields As String 

 

For Each GC In ActiveProject.TaskGroups("Priority Keeping Outline Structure").GroupCriteria 

 If GC.Ascending = True Then 

 Fields = Fields &amp; GC.Index &amp; ". " &amp; GC.FieldName &amp; " is sorted in ascending order." &amp; vbCrLf 

 Else 

 Fields = Fields &amp; GC.Index &amp; ". " &amp; GC.FieldName &amp; " is sorted in descending order." &amp; vbCrLf 

 End If 

Next GC 

 

MsgBox Fields
```

Use the  **[Add](groupcriteria-add-method-project.md)** method to add a **GroupCriterion** object to the **GroupCriteria** collection. The following example adds another criterion to the specified resource group, grouping resources in ascending order as determined by the percentage of their work (in 25-percent increments) that is complete.
 

 



```
ActiveProject.ResourceGroups("Response Pending").GroupCriteria.Add "% Work Complete", True, CellColor:=pjRed, GroupOn:=pjGroupOnPct1_25
```


## Methods



|**Name**|
|:-----|
|[Add](groupcriteria-add-method-project.md)|

## Properties



|**Name**|
|:-----|
|[Application](groupcriteria-application-property-project.md)|
|[Count](groupcriteria-count-property-project.md)|
|[Item](groupcriteria-item-property-project.md)|
|[Parent](groupcriteria-parent-property-project.md)|

## See also


#### Other resources


 
[Project Object Model](http://msdn.microsoft.com/library/900b167b-88ec-ea88-15b7-27bb90c22ac6%28Office.15%29.aspx)
