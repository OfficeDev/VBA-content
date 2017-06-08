---
title: GroupCriterion2 Object (Project)
ms.prod: project-server
api_name:
- Project.GroupCriterion2
ms.assetid: 06047a9d-a9db-43e0-e759-e24560da7128
ms.date: 06/08/2017
---


# GroupCriterion2 Object (Project)

Represents a criterion in a group definition where the group hierarchy can be maintained and cell color can be a hexadecimal value. The  **GroupCriterion2** object is a member of the **[GroupCriteria2](groupcriteria2-object-project.md)** collection.
 


## Example

 **Using the GroupCriterion2 Object**
 

 
Use  **GroupCriteria2(***Index* **)**, where*Index* is the criterion index, to return a single **GroupCriterion2** object. The following example sets the cell color for the first criterion in the Standard Rate resource group to blue.
 

 



```
ActiveProject.ResourceGroups2("Standard Rate").GroupCriteria2(1).CellColor = &amp;HFF0000
```

 **Using the GroupCriteria2 Collection**
 

 
Use the  **[GroupCriteria](group2-groupcriteria-property-project.md)** property to return a **GroupCriteria2** collection. The following example displays a list of the fields used as criteria in the specified task group and shows whether they are sorted in ascending or descending order.
 

 



```
Dim GC2 As GroupCriterion2  
Dim Fields As String  
  
For Each GC2 In ActiveProject.TaskGroups2("Priority Keeping Outline Structure").GroupCriteria  
    If GC2.Ascending = True Then  
        Fields = Fields &amp; GC2.Index &amp; ". " &amp; GC2.FieldName &amp; " is sorted in ascending order." &amp; vbCrLf  
    Else  
        Fields = Fields &amp; GC2.Index &amp; ". " &amp; GC2.FieldName &amp; " is sorted in descending order." &amp; vbCrLf  
    End If  
Next GC2  
  
MsgBox Fields
```

Use the  **[AddEx](groupcriteria2-addex-method-project.md)** method to add a **GroupCriterion2** object to the **GroupCriteria2** collection, where **CellColor** can be a hexadecimal value. The following example adds another criterion to the specified resource group, grouping resources in ascending order as determined by the percentage of their work (in 25-percent increments) that is complete.
 

 



```
ActiveProject.ResourceGroups2("Response Pending").GroupCriteria2.AddEx "% Work Complete", True, _
    CellColor:=&amp;H0101FF, GroupOn:=pjGroupOnPct1_25
```


## Methods



|**Name**|
|:-----|
|[Delete](groupcriterion2-delete-method-project.md)|

## Properties



|**Name**|
|:-----|
|[Application](groupcriterion2-application-property-project.md)|
|[Ascending](groupcriterion2-ascending-property-project.md)|
|[Assignment](groupcriterion2-assignment-property-project.md)|
|[CellColor](groupcriterion2-cellcolor-property-project.md)|
|[CellColorEx](groupcriterion2-cellcolorex-property-project.md)|
|[FieldName](groupcriterion2-fieldname-property-project.md)|
|[FontBold](groupcriterion2-fontbold-property-project.md)|
|[FontColor](groupcriterion2-fontcolor-property-project.md)|
|[FontColorEx](groupcriterion2-fontcolorex-property-project.md)|
|[FontItalic](groupcriterion2-fontitalic-property-project.md)|
|[FontName](groupcriterion2-fontname-property-project.md)|
|[FontSize](groupcriterion2-fontsize-property-project.md)|
|[FontUnderLine](groupcriterion2-fontunderline-property-project.md)|
|[GroupInterval](groupcriterion2-groupinterval-property-project.md)|
|[GroupOn](groupcriterion2-groupon-property-project.md)|
|[Index](groupcriterion2-index-property-project.md)|
|[Parent](groupcriterion2-parent-property-project.md)|
|[Pattern](groupcriterion2-pattern-property-project.md)|
|[StartAt](groupcriterion2-startat-property-project.md)|

## See also


#### Other resources


 
[Project Object Model](http://msdn.microsoft.com/library/900b167b-88ec-ea88-15b7-27bb90c22ac6%28Office.15%29.aspx)
