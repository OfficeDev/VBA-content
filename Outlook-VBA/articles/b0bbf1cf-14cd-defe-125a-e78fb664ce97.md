
# PostItem.Open Event (Outlook)

 **Last modified:** July 28, 2015

Occurs when an instance of the parent object is being opened in an  ** [Inspector](d7384756-669c-0549-1032-c3b864187994.md)**. 

## Syntax

 _expression_. **Open**( **_Cancel_**)

 _expression_A variable that represents a  **PostItem** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
|Cancel|Required| **Boolean**|(Not used in VBScript).  **False** when the event occurs. If the event procedure sets this argument to **True**, the open operation is not completed and the inspector is not displayed.|

## Remarks

When this event occurs, the  **Inspector** object is initialized but not yet displayed. The **Open** event differs from the ** [Read](aa39ec06-19ed-4655-6990-e4c4c45649d5.md)** event in that **Read** occurs whenever the user selects the item in a view that supports in-cell editing as well as when the item is being opened in an inspector.

In Microsoft Visual Basic Scripting Edition (VBScript), if you set the return value of this function to  **False**, the open operation is not completed and the inspector is not displayed.


## See also


#### Concepts


 [PostItem Object](de44065d-4e93-315a-279f-7b92f09c0465.md)
#### Other resources


 [PostItem Object Members](5b150db1-c96d-0721-ec36-d5b5ebc20fd8.md)
