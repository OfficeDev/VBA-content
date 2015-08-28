
# DataLabels.Propagate Method (Excel)

 **Last modified:** July 28, 2015

Enables you to take the contents and formatting of a single data label and apply it to every other data label on the series.

## Syntax

 _expression_. **Propagate**(Index)

 _expression_A variable that represents a  **DataLabels** object.


### Parameters



|**Name**|**Required/Optional**|**Data type**|**Description**|
|:-----|:-----|:-----|:-----|
|Index|Required|VARIANT|The index number of the data label to propagate.|

### Remarks

If the source data label supports fields that are incompatible with the destination data label, those fields will be inserted as their [FIELD NAME]. For example, if a data label on a pie series with a percentage field is propagated to a data label on a column chart, that percentage field will become an unresolved field showing [PERCENTAGE].


 **Note**  Passing an argument of zero (0) resets the data labels to your current prototype.


### Return value

 **VOID**


## See also


#### Concepts


 [DataLabels Object](3d79271e-c702-e785-6984-d838d060a8c5.md)
#### Other resources


 [DataLabels Object Members](3c9d909d-d090-b6ed-8a28-ba62c3459044.md)
