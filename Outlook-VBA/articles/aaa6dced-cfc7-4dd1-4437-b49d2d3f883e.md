
# DistListItem.Close Event (Outlook)

 **Last modified:** July 28, 2015

Occurs when the inspector associated with an item (which is an instance of the parent object) is being closed.

## Syntax

 _expression_. **Close**( **_Cancel_**)

 _expression_A variable that represents a  **DistListItem** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
|Cancel|Required| **Boolean**|(Not used in VBScript).  **False** when the event occurs. If the event procedure sets this argument to **True**, the close operation is not completed and the inspector is left open.|

## Remarks

In Microsoft Visual Basic Scripting Edition (VBScript), if you set the return value of this function to  **False**, the close operation isn't completed and the inspector is left open.

If you use the  ** [Close](6e56d716-ec8b-4a4c-1b1a-061f659f5c08.md)**method to fire this event, it can only be canceled if the  **Close** method uses the **olPromptForSave** argument.


## See also


#### Concepts


 [DistListItem Object](027c3986-abff-d9b1-ecc2-26d60805e952.md)
#### Other resources


 [DistListItem Object Members](3ba4af84-ce84-61d9-1bc9-fab41bf6f125.md)
