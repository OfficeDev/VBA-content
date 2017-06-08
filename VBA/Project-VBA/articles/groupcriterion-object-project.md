---
title: GroupCriterion Object (Project)
ms.prod: project-server
api_name:
- Project.GroupCriterion
ms.assetid: 9c3f7a79-c65f-925c-98ae-c217bd6ed8f7
ms.date: 06/08/2017
---


# GroupCriterion Object (Project)

Represents a criterion in a group definition. The  **GroupCriterion** object is a member of the **[GroupCriteria](groupcriteria-object-project.md)** collection.
 


## Remarks

To use groups where the group hierarchy can be maintained and cell color can be a hexadecimal value, see the  **[GroupCriterion2](groupcriterion2-object-project.md)** object.
 

 

## Example

 **Using the GroupCriterion Object**
 

 
Use  **GroupCriteria(***Index* **)**, where*Index* is the criterion index, to return a single **GroupCriterion** object. The following example sets the cell color for the first criterion in the Standard Rate resource group to blue.
 

 



```
ActiveProject.ResourceGroups("Standard Rate").GroupCriteria(1).CellColor = pjBlue
```

 **Using the GroupCriteria Collection**
 

 
Use the  **[GroupCriteria](group-groupcriteria-property-project.md)** property to return a **GroupCriteria** collection. The following example displays a list of the fields used as criteria in the specified task group and shows whether they are sorted in ascending or descending order.
 

 



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
|[Delete](groupcriterion-delete-method-project.md)|

## Properties



|**Name**|
|:-----|
|[Application](groupcriterion-application-property-project.md)|
|[Ascending](groupcriterion-ascending-property-project.md)|
|[Assignment](groupcriterion-assignment-property-project.md)|
|[CellColor](groupcriterion-cellcolor-property-project.md)|
|[FieldName](groupcriterion-fieldname-property-project.md)|
|[FontBold](groupcriterion-fontbold-property-project.md)|
|[FontColor](groupcriterion-fontcolor-property-project.md)|
|[FontItalic](groupcriterion-fontitalic-property-project.md)|
|[FontName](groupcriterion-fontname-property-project.md)|
|[FontSize](groupcriterion-fontsize-property-project.md)|
|[FontUnderLine](groupcriterion-fontunderline-property-project.md)|
|[GroupInterval](groupcriterion-groupinterval-property-project.md)|
|[GroupOn](groupcriterion-groupon-property-project.md)|
|[Index](groupcriterion-index-property-project.md)|
|[Parent](groupcriterion-parent-property-project.md)|
|[Pattern](groupcriterion-pattern-property-project.md)|
|[StartAt](groupcriterion-startat-property-project.md)|

