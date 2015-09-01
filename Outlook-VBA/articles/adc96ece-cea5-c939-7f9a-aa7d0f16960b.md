
# TaskItem.PropertyChange Event (Outlook)

 **Last modified:** July 28, 2015

Occurs when an explicit built-in property (for example,  ** [Subject](9f487fbc-48ab-e01d-c1a4-5b67fcb1a118.md)**) of an instance of the parent object is changed. 

## Syntax

 _expression_. **PropertyChange**( **_Name_**)

 _expression_A variable that represents a  **TaskItem** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
|Name|Required| **String**|The name of the property that was changed.|

## Remarks

The property name is passed to the event so that you can determine which property was changed.


## See also


#### Concepts


 [TaskItem Object](5df8cfa5-5460-a5a1-a130-ba5bca1a0091.md)
#### Other resources


 [TaskItem Object Members](97234a76-2fc5-bbe4-2e14-25ae18694fc9.md)
