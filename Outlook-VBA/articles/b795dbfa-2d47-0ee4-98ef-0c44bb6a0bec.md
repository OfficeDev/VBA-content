
# SharingItem.Open Event (Outlook)

 **Last modified:** July 28, 2015

Occurs when an instance of the parent object is being opened in an  ** [Inspector](d7384756-669c-0549-1032-c3b864187994.md)**. 

## Syntax

 _expression_. **Open**( **_Cancel_**)

 _expression_An expression that returns a  **SharingItem** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
|Cancel|Required| **Boolean**|(Not used in VBScript).  **False** when the event occurs. If the event procedure sets this argument to **True**, the open operation is not completed and the inspector is not displayed.|

## Remarks

When this event occurs, the  **Inspector** object is initialized but not yet displayed. The **Open** event differs from the ** [Read](2bcf07e6-e9c1-b3ce-118c-a2c82b48ff5f.md)** event in that **Read** occurs whenever the user selects the item in a view that supports in-cell editing as well as when the item is being opened in an inspector.

In Microsoft Visual Basic Scripting Edition (VBScript), if you set the return value of this function to  **False**, the open operation is not completed and the inspector is not displayed.


## See also


#### Concepts


 [SharingItem Object](63dd3451-44f3-7cc4-c6e2-7dad5835a7d2.md)
#### Other resources


 [SharingItem Object Members](719ad60e-2242-2c54-778f-006b61690389.md)
