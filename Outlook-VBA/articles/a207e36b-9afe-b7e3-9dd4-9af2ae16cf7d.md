
# OlkCheckBox.AfterUpdate Event (Outlook)

 **Last modified:** July 28, 2015

Occurs after the data in the control has been changed through the user interface.

## Syntax

 _expression_. **AfterUpdate**

 _expression_A variable that represents an  **OlkCheckBox** object.


## Remarks

 **BeforeUpdate** and **AfterUpdate** can occur any time the data in the control is being saved to the item. The typical sequence of events involving **AfterUpdate** for this control is as follows:


1. User focuses on the control
    
2.  **BeforeUpdate**
    
3. Control data is updated
    
4.  ** AfterUpdate**
    
5.  **Exit**: User moves focus away from control
    



## See also


#### Concepts


 [OlkCheckBox Object](79460205-a604-7011-a9b3-14e651807f09.md)
#### Other resources


 [OlkCheckBox Object Members](acf62b06-215d-6b2b-57b0-ccbfd0c92aed.md)
