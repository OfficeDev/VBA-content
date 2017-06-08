---
title: GroupCriteria2 Object (Project)
ms.prod: project-server
ms.assetid: ac785cc4-dbe3-0b1d-d1f1-6d45c93bfb1d
ms.date: 06/08/2017
---


# GroupCriteria2 Object (Project)

Contains a collection of  **[GroupCriterion2](groupcriterion2-object-project.md)** objects, where the group hierarchy can be maintained and cell color can be a hexadecimal value.
 


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
       Fields = Fields &amp; GC2.Index &amp; ". " &amp; GC2.FieldName &amp; " is sorted in ascending order." _
           &amp; vbCrLf  
    Else  
        Fields = Fields &amp; GC2.Index &amp; ". " &amp; GC2.FieldName &amp; " is sorted in descending order." _
           &amp; vbCrLf  
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
|[Add](groupcriteria2-add-method-project.md)|
|[AddEx](groupcriteria2-addex-method-project.md)|

## Properties



|**Name**|
|:-----|
|[Application](groupcriteria2-application-property-project.md)|
|[Count](groupcriteria2-count-property-project.md)|
|[Item](groupcriteria2-item-property-project.md)|
|[Parent](groupcriteria2-parent-property-project.md)|

## See also


#### Other resources


 
[Project Object Model](http://msdn.microsoft.com/library/900b167b-88ec-ea88-15b7-27bb90c22ac6%28Office.15%29.aspx)
