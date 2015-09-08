
# Form.Close Event (Access)

 **Last modified:** July 28, 2015

The  **Close** event occurs when a form is closed and removed from the screen.

## Syntax

 _expression_. **Close**

 _expression_A variable that represents a  **Form** object.


### Return Value

nothing


## Remarks

To run a macro or event procedure when this event occurs, set the  **OnClose**property to the name of the macro or to [Event Procedure].

The  **Close** event occurs after the **Unload**event, which is triggered after the form is closed but before it is removed from the screen.

When you close a form, the following events occur in this order:

 **Unload** â†’ **Deactivate** â†’ **Close**

When the  **Close** event occurs, you can open another window or request the user's name to make a log entry indicating who used the form or report.

The  **Unload** event can be canceled, but the **Close** event can't.


## See also


#### Concepts


 [Form Object](72ef9219-142b-b690-b696-3eba9a5d4522.md)
#### Other resources


 [Form Object Members](e1976b58-28ca-8f76-cdf3-6732cb06ce6c.md)
