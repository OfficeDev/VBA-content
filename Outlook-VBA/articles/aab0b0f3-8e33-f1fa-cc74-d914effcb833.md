
# ReportItem.Send Event (Outlook)

 **Last modified:** July 28, 2015

Occurs when the user selects the  **Send** action for an item (which is an instance of the parent object).

## Syntax

 _expression_. **Send**( **_Cancel_**)

 _expression_A variable that represents a  **ReportItem** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
|Cancel|Required| **Boolean**|(Not used in VBScript).  **False** when the event occurs. If the event procedure sets this argument to **True**, the send operation is not completed and the inspector is left open.|

## Remarks

In Microsoft Visual Basic Scripting Edition (VBScript), if you set the return value of this function to  **False**, the item is not sent.


## See also


#### Concepts


 [ReportItem Object](16ebe336-72e0-42f6-99d3-edecc3ea284d.md)
#### Other resources


 [ReportItem Object Members](5a5662dd-e969-bbd5-129b-44609ba1cf9f.md)
