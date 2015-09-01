
# PostItem.AfterWrite Event (Outlook)

 **Last modified:** July 28, 2015

Occurs after Microsoft Outlook has saved the item.

## Syntax

 _expression_. **AfterWrite**

 _expression_A variable that represents a  **PostItem** object.


## Remarks

The  **AfterWrite** event occurs after the ** [Write](27ab5442-2ce2-c40e-b95c-6e23f29e124b.md)** event. This event is not cancelable. To determine when the item is unloaded from memory, use the ** [Unload](42dea931-c3dd-b8ff-5ace-0744b17650e0.md)** event.

The  **AfterWrite** event corresponds to the Exchange Client Extensions (ECE) event **IExchExtMessageEvents::OnWriteComplete**.

Only the following members of the item object can be accessed in the  **AfterWrite** event:


-  ** [Class](a79d30fc-04a1-36cb-c4f5-8c9cc063b89e.md)**
    
-  ** [MessageClass](4f5064a7-0de0-025b-56f9-3c29c4741e5a.md)**
    
-  **MAPIOBJECT**
    
The  **MAPIOBJECT** property is a hidden property in the Outlook object model. This property provides access to the underlying MAPI ** [IMessage](http://msdn.microsoft.com/en-us/library/cc842097%28office.14%29.aspx)** object, and can be invoked only via the ** [IUnknown](http://msdn.microsoft.com/en-us/library/ms680509%28VS.85%29.aspx)** interface. The property is accessible to programs written in languages such as C or C++ that support **IUnknown**.  **MAPIOBJECT** is not available through the ** [IDispatch](http://msdn.microsoft.com/en-us/library/ms221608.aspx)** interface. Development languages such as Visual Basic for Applications (VBA), Visual C#, and Visual Basic support the **IDispatch** interface and not **IUnknown**, and therefore, they cannot access  **MAPIOBJECT**. If other properties or methods of the parent item are accessed in this event, Outlook raises an error.

The object obtained from the  **MAPIOBJECT** property in this event must contain all the changes persisted by Outlook. The implementer can call the ** [SaveChanges](http://msdn.microsoft.com/en-us/library/cc842181%28office.14%29.aspx)** method on the **IMessage** object to persist changes to the underlying **IMessage** object represented by **MAPIOBJECT**, and Outlook will not revert those changes.

Implementers must release the object obtained from the  **MAPIOBJECT** property in the event before the event completes. Attempting to use that object outside the context of the event is unsupported and will lead to unpredictable behavior.


## See also


#### Concepts


 [PostItem Object](de44065d-4e93-315a-279f-7b92f09c0465.md)
#### Other resources


 [PostItem Object Members](5b150db1-c96d-0721-ec36-d5b5ebc20fd8.md)
