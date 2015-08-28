
# Application.SelectBeginning Method (Project)

 **Last modified:** July 28, 2015

 **In this article**
 [Syntax](#sectionSection0)
 [Remarks](#sectionSection1)
 [Example](#sectionSection2)


Selects the first cell in the active table or view.


## Syntax
<a name="sectionSection0"> </a>

 _expression_. **SelectBeginning**( **_Extend_**)

 _expression_A variable that represents an  **Application** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
|Extend|Optional| **Boolean**| **True** if the current selection is extended to the first cell. If the active view is the Network Diagram or Resource Graph, Extend is ignored. The default value is **False**.|

### Return Value

 **Boolean**


## Remarks
<a name="sectionSection1"> </a>

In the Resource Graph,  **SelectBeginning** selects the resource with the lowest identification number. In the Network Diagram, **SelectBeginning** selects the box closest to the upper-left corner of the view.


## Example
<a name="sectionSection2"> </a>

The following example selects the "Name" field of row 4 as the beginning field in the Gantt Chart.


```
Sub Select_Beginning() 
 
 ViewApply Name:="&amp;Gantt Chart" 
 SelectTaskField Row:=4, Column:="Name", RowRelative:=False 
 
 SelectBeginning Extend:=True 
End Sub
```

