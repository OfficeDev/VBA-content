
# onReadyStateChange Event (RDS)

 **Last modified:** June 29, 2011

 _ **Applies to:** Access 2013 | Access 2016_

 **In this article**
[Syntax](#sectionSection1)
[Parameters](#sectionSection2)
[Remarks](#sectionSection3)



The  **onReadyStateChange** event is called whenever the value of the[ReadyState](e7b62205-a604-ef43-2f5d-9b51b46d2b5a.md) property changes.

## Syntax
<a name="sectionSection1"> </a>

 **onReadyStateChange**


## Parameters
<a name="sectionSection2"> </a>

None.


## Remarks
<a name="sectionSection3"> </a>

The  **ReadyState** property reflects the progress of an[RDS.DataControl](ac430669-7628-696c-c036-b5d35405d788.md) object as it asynchronously retrieves data into its[Recordset](0f963bf8-f066-dc8a-b754-f427de712df1.md) object. Use the **onReadyStateChange** event to monitor changes in the **ReadyState** property whenever they occur. This is more efficient than periodically checking the property's value.

