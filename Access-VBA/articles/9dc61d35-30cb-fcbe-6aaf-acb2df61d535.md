
# IRtdServer.Heartbeat Method (Excel)

 **Last modified:** July 28, 2015

Determines if the real-time data server is still active. Returns a  **Long** value. Zero or a negative number indicates failure; a positive number indicates that the server is active.

## Syntax

 _expression_. **Heartbeat**

 _expression_A variable that represents an  **IRtdServer** object.


### Return Value

Long


## Remarks

The  **Heartbeat** method is called by Microsoft Excel if the ** [HeartbeatInterval](45a3df85-59c1-fedb-e94b-8f011601fc72.md)**property has elapsed since the last time Excel was called with the  ** [UpdateNotify](e3ae5a7e-4d8c-9eba-62ab-a24d1045bc77.md)**method.


## See also


#### Concepts


 [IRtdServer Object](6a85aa64-9514-74bb-3c63-141275f1b671.md)
#### Other resources


 [IRtdServer Object Members](90baa971-8dc0-b4b9-77c4-72530f1aaf21.md)
