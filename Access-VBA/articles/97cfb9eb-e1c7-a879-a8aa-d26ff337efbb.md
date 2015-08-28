
# Form.OnDelete Property (Access)

 **Last modified:** July 28, 2015

 **In this article**
 [Syntax](#sectionSection0)
 [Remarks](#sectionSection1)
 [Example](#sectionSection2)


Sets or returns the value of the  **On Delete** box in the **Properties** window of a form. Read/write **String**.


## Syntax
<a name="sectionSection0"> </a>

 _expression_. **OnDelete**

 _expression_A variable that represents a  **Form** object.


## Remarks
<a name="sectionSection1"> </a>

This property is helpful for programmatically changing the action Microsoft Access takes when an event is triggered. For example, between event calls you may want to change an expression's parameters, or switch from an event procedure to an expression or macro, depending on the circumstances under which the event was triggered. 

The  **Delete** event occurs when the user performs some action, such as pressing the DELETE key to delete a record, but before the record is actually deleted.

The  **OnDelete** value will be one of the following, depending on the selection chosen in the **Choose Builder** window (accessed by clicking the **Build** button next to the **On Delete** box in the form's **Properties** window):


- If Expression Builder is chosen, the value will be "= _expression_", where  _expression_ is the expression from the Expression Builder window.
    
- If Macro Builder is chosen, the value is the name of the macro. 
    
- If Code Builder is chosen, the value will be "[Event Procedure]". 
    
If the  **On Delete** box is blank, the property value is an empty string.


## Example
<a name="sectionSection2"> </a>

The following example associates the  **Delete** event with the "Form_Delete" event for the "Order Entry" form.


```
Forms("Order Entry").OnDelete = "[Event Procedure]"
```


## See also
<a name="sectionSection2"> </a>


#### Concepts


 [Form Object](72ef9219-142b-b690-b696-3eba9a5d4522.md)
#### Other resources


 [Form Object Members](e1976b58-28ca-8f76-cdf3-6732cb06ce6c.md)
