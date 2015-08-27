
# Event.TargetArgs Property (Visio)

 **Last modified:** July 28, 2015

 _**Applies to:** Visio 2013 Preview_

Gets or sets the arguments to be sent to the target of an event. Read/write.


## Syntax

 _expression_. **TargetArgs**

 _expression_A variable that represents a  **Event** object.


### Return Value

String


## Remarks

An event consists of an event-action pair. When the event occurs, the action is performed. An event also specifies the target of the action and arguments to send to the target.

When you use  **visActCodeRunAddon**, the  **TargetArgs** property contains the arguments to send to the add-on when it is run.

When you use  **visActCodeAdvise**, the  **TargetArgs** property contains the string specified with the **AddAdvise** method when the **Event** object was created. When the program receives notification of the event, it can get the **Event** object and its **TargetArgs** property to obtain the string.

