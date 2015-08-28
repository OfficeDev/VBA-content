
# PivotFilters.Add2 Method (Excel)

 **Last modified:** July 28, 2015

Adds new filters to the  **PivotFilters** collection.

## Syntax

 _expression_. **Add2**( **_Type_**,  **_DataField_**,  **_Value1_**,  **_Value2_**,  **_Order_**,  **_Name_**,  **_Description_**,  **_MemberPropertyField_**,  **_WholeDayFilter_**)

 _expression_A variable that represents a  **PivotFilters** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
|Type|Required| **XlPivotFilterType**|Requires an  ** [XlPivotFilterType](0ae3f0fe-02e3-b0f7-1506-1961c4adcd6c.md)** type of filter.|
|DataField|Optional| **Variant**|The field to which the filter is attached.|
|Value1|Optional| **Variant**|Filter value 1.|
|Value2|Optional| **Variant**|Filter value 2.|
|Order|Optional| **Variant**|Order in which the data should be filtered.|
|Name|Optional| **Variant**|Name of the filter.|
|Description|Optional| **Variant**|A brief description of the filter.|
|MemberPropertyField|Optional| **Variant**|Specifies the member property field on which the label filter is based.|
|WholeDayFilter|Optional| **Variant**|Specifies a filter based on days.|

### Return Value

PivotFilter


## Example

Following are some examples of how to use the  **Add** function correctly.


```
ActiveCell.PivotField.PivotFilters.Add FilterType := xlThisWeek 
 
ActiveCell.PivotField.PivotFilters.Add FilterType := xlTopCount DataField := MyPivotField2 Value1 := 10 
 
ActiveCell.PivotField.PivotFilters.Add FilterType := xlCaptionIsNotBetween Value1 := "A" Value2 := "G" 
 
ActiveCell.PivotField.PivotFilters.Add FilterType := xlValueIsGreaterThanOrEqualTo DataField := MyPivotField2 Value1 := 10000 

```

The following example returns a run-time error because the data type of Value1 is invalid.




```
ActiveCell.PivotField.PivotFilters.Add FilterType := xlValueIsGreaterThanOrEqualTo DataField := MyPivotField2 Value1 := "Allan"
```


## See also


#### Concepts


 [PivotFilters Object](fc647acb-bd6a-8544-6411-1f5e49807e53.md)
#### Other resources


 [PivotFilters Object Members](57f1f375-1b7b-c488-c236-91ed26a68bb6.md)
