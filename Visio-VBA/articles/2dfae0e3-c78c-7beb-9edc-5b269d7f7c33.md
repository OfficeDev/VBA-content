
# MSGWrap.ptx Property (Visio)

 **Last modified:** July 28, 2015

 _**Applies to:** Visio 2013 Preview_

Gets or sets the  **pt.x** member of the **MSG** structure being wrapped. Read/write.


## Syntax

 _expression_. **ptx**

 _expression_A variable that represents a  **MSGWrap** object.


### Return Value

Long


## Remarks

The  **ptx** property corresponds to the **pt.x** member in the **MSG** structure defined as part of the Microsoft Windows operating system. If an event handler is handling the **OnKeystrokeMessageForAddon** event, Microsoft Visio passes a **MSGWrap** object as an argument when this event fires. A **MSGWrap** object is a wrapper around the Windows **MSG** structure.

For details, search for "MSG structure" on MSDN, the Microsoft Developer Network.

