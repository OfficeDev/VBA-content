
# ContactItem.BeforeRead Event (Outlook)

 **Last modified:** July 28, 2015

Occurs before Microsoft Outlook begins to read the properties for the item.

## Syntax

 _expression_. **BeforeRead**

 _expression_A variable that represents a  **ContactItem** object.


## Remarks

The  **BeforeRead** event occurs before the ** [Read](508b4637-9d74-7645-7719-3c148d0688d8.md)** event. Unlike other events with the Before prefix, this event is not cancelable. To determine when the item is unloaded from memory, use the ** [Unload](16a3d7ce-0843-5eb5-bbea-df6557ceda05.md)** event.

The  **BeforeRead** event corresponds to the Exchange Client Extensions (ECE) event **IExchExtMessageEvents::OnRead**.

Only the following members of the item object can be accessed in the  **BeforeRead** event:


-  ** [Class](7c08cb72-fdbb-aac8-2691-382bfdae22c8.md)**
    
-  ** [MessageClass](3d6594b7-8abe-9e49-64e0-be3062807e34.md)**
    
-  **MAPIOBJECT**
    
The  **MAPIOBJECT** property is a hidden property in the Outlook object model. This property provides access to the underlying MAPI ** [IMessage](http://msdn.microsoft.com/en-us/library/cc842097%28office.14%29.aspx)** object, and can be invoked only via the ** [IUnknown](http://msdn.microsoft.com/en-us/library/ms680509%28VS.85%29.aspx)** interface. The property is accessible to programs written in languages such as C or C++ that support **IUnknown**.  **MAPIOBJECT** is not available through the ** [IDispatch](http://msdn.microsoft.com/en-us/library/ms221608.aspx)** interface. Development languages such as Visual Basic for Applications (VBA), Visual C#, and Visual Basic support the **IDispatch** interface and not **IUnknown**, and therefore, they cannot access  **MAPIOBJECT**. If other properties or methods of the parent item are accessed in this event, Outlook raises an error.

If the implementer accesses the underlying  **IMessage** object and changes properties on that object, Outlook will render that item reflecting the changes to the **IMessage** object. The implementer does not have to call ** [SaveChanges](http://msdn.microsoft.com/en-us/library/cc842181%28office.14%29.aspx)** on the **IMessage** object to cause the changes to be reflected in Outlook.

Implementers must release the object obtained from the  **MAPIOBJECT** property in the event before the event completes. Attempting to use that object outside the context of the event is unsupported and will lead to unpredictable behavior.


## See also


#### Concepts


 [ContactItem Object](8e32093c-a678-f1fd-3f35-c2d8994d166f.md)
#### Other resources


 [ContactItem Object Members](a8b13369-4c87-02aa-e62a-1f3067e559fa.md)
