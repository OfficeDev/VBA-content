
# PostItem.Forward Event (Outlook)

 **Last modified:** July 28, 2015

Occurs when the user selects the  **Forward** action for an item, or when the **Forward** method is called for the item, which is an instance of the parent object.

## Syntax

 _expression_. **Forward**( **_Forward_**,  **_Cancel_**)

 _expression_A variable that represents a  **PostItem** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
|Forward|Required| **Object**|The new item being forwarded.|
|Cancel|Required| **Boolean**|(Not used in VBScript).  **False** when the event occurs. If the event procedure sets this argument to **True**, the forward operation is not completed and the new item is not displayed.|

## Remarks

In VBScript, if you set the return value of this function to  **False**, the forward action is not completed and the new item is not displayed.


## See also


#### Concepts


 [PostItem Object](de44065d-4e93-315a-279f-7b92f09c0465.md)
#### Other resources


 [PostItem Object Members](5b150db1-c96d-0721-ec36-d5b5ebc20fd8.md)
