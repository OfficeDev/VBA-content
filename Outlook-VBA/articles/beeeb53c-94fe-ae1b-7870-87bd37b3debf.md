
# ContactItem.Close Event (Outlook)

 **Last modified:** July 28, 2015

Occurs when the inspector associated with an item (which is an instance of the parent object) is being closed.

## Syntax

 _expression_. **Close**( **_Cancel_**)

 _expression_A variable that represents a  **ContactItem** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
|Cancel|Required| **Boolean**|(Not used in VBScript).  **False** when the event occurs. If the event procedure sets this argument to **True**, the close operation is not completed and the inspector is left open.|

## Remarks

In Microsoft Visual Basic Scripting Edition (VBScript), if you set the return value of this function to  **False**, the close operation isn't completed and the inspector is left open. 

If you use the  ** [Close](17cd04b5-1bf1-5df1-b1f4-f6e488d00fd5.md)** method to fire this event, it can only be canceled if the **Close** method uses the **olPromptForSave** argument.


## See also


#### Concepts


 [ContactItem Object](8e32093c-a678-f1fd-3f35-c2d8994d166f.md)
#### Other resources


 [ContactItem Object Members](a8b13369-4c87-02aa-e62a-1f3067e559fa.md)
