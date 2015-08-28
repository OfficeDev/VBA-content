
# Application.SetAutoFilter Method (Project)

 **Last modified:** July 28, 2015

 **In this article**
 [Syntax](#sectionSection0)
 [Remarks](#sectionSection1)
 [Example](#sectionSection2)


Sets the criteria for an AutoFilter for a specified field in a sheet view.


## Syntax
<a name="sectionSection0"> </a>

 _expression_. **SetAutoFilter**( **_FieldName_**,  **_FilterType_**,  **_Test1_**,  **_Criteria1_**,  **_Operation_**,  **_Test2_**,  **_Criteria2_**)

 _expression_An expression that returns an  **Application** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
|FieldName|Required| **String**|Name of the field.|
|FilterType|Optional| **PjAutoFilterType**|Type of filter; can be one of the  ** [PjAutoFilterType](f7bd2ed9-90a1-63e9-493c-28c9c944795b.md)** constants. The default value is **pjAutoFilterClear**, which clears the AutoFilter. |
|Test1|Optional| **String**|Specifies the type of comparison for the first test. Requires that FilterType is **pjAutoFilterCustom**, and that Criteria1 specifies a value. Can be one of the following comparison strings:

|**Comparison String**|**Description**|
|:-----|:-----|
|"equals"| The value of _FieldName_ equals _Criteria1_. |
|"does not equal"| The value of _FieldName_ does not equal _Criteria1_. |
|"is greater than"| The value of _FieldName_ is greater than _Criteria1_. |
|"is greater than or equal to"| The value of _FieldName_ is greater than or equal to _Criteria1_. |
|"is less than"| The value of _FieldName_ is less than _Criteria1_. |
|"is less than or equal to"| The value of _FieldName_ is less than or equal to _Criteria1_. |
|"is within"| The value of _FieldName_ is within _Criteria1_. |
|"is not within"| The value of _FieldName_ is not within _Criteria1_. |
|
|Criteria1|Optional| **String**|The value of the first comparison with the value of the field specified by FieldName.|
|Operation|Optional| **String**|The logical operation if there is a second test. The Operation value can be "And" or "Or".|
|Test2|Optional| **String**|Specifies the type of comparison for the second test. Requires that FilterType is **pjAutoFilterCustom**, the Operation value must be set, and thatCriteria2 specifies a value. The string can be one of the comparisons in the table for Test1.|
|Criteria2|Optional| **String**|The value of the second comparison with the value of the field specified by FieldName.|

### Return Value

 **Boolean**


## Remarks
<a name="sectionSection1"> </a>

To turn the AutoFilter feature on or off, see the  ** [AutoFilter](391d5a61-cba3-9e28-c448-d0befcc456c7.md)** method.


 **Note**  A column name in a sheet view can have a different title than the name of the field it shows.


## Example
<a name="sectionSection2"> </a>

The following example sets a custom AutoFilter for the "% Work Complete" field. 


```
Sub TestAutoFilter() 
    If Not ActiveProject.AutoFilter Then 
        Application.AutoFilter 
    End If 
 
    Application.SetAutoFilter FieldName:="% Work Complete", FilterType:=pjAutoFilterCustom, _ 
    Test1:="equals", Criteria1:="0%" 
End Sub
```

If there is an AutoFilter set for the "% Work Complete" field, the following line of code clears the AutoFilter because the default value for the optional FilterType argument is **pjAutoFilterClear**.




```
Application.SetAutoFilter FieldName:="% Work Complete"
```

