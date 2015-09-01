
# OlkListBox.AfterUpdate Event (Outlook)

 **Last modified:** July 28, 2015

Occurs after the data in the control has been changed through the user interface.

## Syntax

 _expression_. **AfterUpdate**

 _expression_A variable that represents an  **OlkListBox** object.


## Remarks

 **BeforeUpdate** and **AfterUpdate** can occur any time the data in the control is being saved to the item. The typical sequence of events involving **AfterUpdate** for this control is as follows:


1. User focuses on the control
    
2.  **BeforeUpdate**
    
3. Control data is updated
    
4.  ** AfterUpdate**
    
5.  **Exit**: User moves focus away from control
    



## See also


#### Concepts


 [OlkListBox Object](373d2a00-97e5-2ed3-f15f-577d97b32334.md)
#### Other resources


 [OlkListBox Object Members](b8bed0b5-6994-1492-055e-4067b232f9c4.md)
