
# Chart.SeriesCollection Method (Excel)

 **Last modified:** July 28, 2015

Returns an object that represents either a single series (a  ** [Series](c7d34b32-8172-f7a0-0a17-f01d44246b64.md)** object) or a collection of all the series (a ** [SeriesCollection](93aa1f0b-4939-8c60-a444-2f791e8ce144.md)**collection) in the chart or chart group.

## Syntax

 _expression_. **SeriesCollection**( **_Index_**)

 _expression_A variable that represents a  **Chart** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
|Index|Optional| **Variant**|The name or number of the series.|

### Return Value

Object


## Example

This example turns on data labels for series one in Chart1.


```
Charts("Chart1").SeriesCollection(1).HasDataLabels = True
```


## See also


#### Concepts


 [Chart Object](179c32ce-49bd-6f36-ea12-89fb5443f3ea.md)
#### Other resources


 [Chart Object Members](a3f8ac44-02d6-6f3f-b5e0-23f4bd5d6baf.md)
