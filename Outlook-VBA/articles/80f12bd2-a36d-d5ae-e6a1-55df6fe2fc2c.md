
# ContactItem.Open Event (Outlook)

 **Last modified:** July 28, 2015

Occurs when an instance of the parent object is being opened in an  ** [Inspector](d7384756-669c-0549-1032-c3b864187994.md)**. 

## Syntax

 _expression_. **Open**( **_Cancel_**)

 _expression_A variable that represents a  **ContactItem** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
|Cancel|Required| **Boolean**|(Not used in VBScript).  **False** when the event occurs. If the event procedure sets this argument to **True**, the open operation is not completed and the inspector is not displayed.|

## Remarks

When this event occurs, the  **Inspector** object is initialized but not yet displayed. The **Open** event differs from the ** [Read](aa39ec06-19ed-4655-6990-e4c4c45649d5.md)** event in that **Read** occurs whenever the user selects the item in a view that supports in-cell editing as well as when the item is being opened in an inspector.

In Microsoft Visual Basic Scripting Edition (VBScript), if you set the return value of this function to  **False**, the open operation is not completed and the inspector is not displayed.


## See also


#### Concepts


 [ContactItem Object](8e32093c-a678-f1fd-3f35-c2d8994d166f.md)
#### Other resources


 [ContactItem Object Members](a8b13369-4c87-02aa-e62a-1f3067e559fa.md)
