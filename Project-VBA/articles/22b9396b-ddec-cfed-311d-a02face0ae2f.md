
# Application.SelectResourceColumn Method (Project)

 **Last modified:** July 28, 2015

 **In this article**
 [Syntax](#sectionSection0)
 [Remarks](#sectionSection1)
 [Example](#sectionSection2)


Selects a column containing resource information.


## Syntax
<a name="sectionSection0"> </a>

 _expression_. **SelectResourceColumn**( **_Column_**,  **_Additional_**,  **_Extend_**,  **_Add_**)

 _expression_A variable that represents an  **Application** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
|Column|Optional| **String**|The field name of the column to select. The default is the column containing the active cell.|
|Additional|Optional| **Integer**|The number of additional columns to select to the right of  **Column**. If  **Extend** is **True**,  **Additional** is ignored. The default value is 0.|
|Extend|Optional| **Boolean**| **True** if all columns between the current selection and **Column** are selected. The default value is **False**.|
|Add|Optional| **Boolean**| **True** if the current column is included in the selection. The default value is **False**.|

### Return Value

 **Boolean**


## Remarks
<a name="sectionSection1"> </a>

The  **SelectResourceColumn** method is only available when the Resource Sheet or Resource Usage view is the active view.


## Example
<a name="sectionSection2"> </a>

The following example selects the  **Indicators** column and the next two columns.


```
Sub Select_ResourceColumn() 
 
 'Activate Resource Sheet 
 ViewApply Name:="&amp;Resource Sheet" 
 SelectResourceColumn Column:="Indicators", Additional:=2 
End Sub
```

