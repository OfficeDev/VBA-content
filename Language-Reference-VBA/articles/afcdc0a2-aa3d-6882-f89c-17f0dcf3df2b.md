
# CreateEventProc Method (VBA Add-In Object Model)

 **Last modified:** July 28, 2015


Creates an event  [procedure](b8bdf64f-5920-1ae9-16d0-b26d09524a30.md).
 **Syntax**
 _object_**.CreateEventProc(**_eventname_,  _objectname_**) As Long**
The  **CreateEventProc** syntax has these parts:


|**Part**|**Description**|
|:-----|:-----|
| _object_|Required. An  [object expression](b8bdf64f-5920-1ae9-16d0-b26d09524a30.md) that evaluates to an object in the Applies To list.|
| _eventname_|Required. A  [string expression](b8bdf64f-5920-1ae9-16d0-b26d09524a30.md) specifying the name of the event you want to add to the [module](b8bdf64f-5920-1ae9-16d0-b26d09524a30.md).|
| _objectname_|Required. A string expression specifying the name of the object that is the source of the event.|
 **Remarks**
Use the  **CreateEventProc** method to create an event procedure. For example, to create an event procedure for the **Click** event of a **Command Button** control named `Command1` you would use the following code, where `CM` represents an object of type **CodeModule**:



```
TextLocation = CM.CreateEventProc("Click", "Command1")
```

The  **CreateEventProc** method returns the line at which the body of the event procedure starts. **CreateEventProc** fails if the [arguments](b8bdf64f-5920-1ae9-16d0-b26d09524a30.md) refer to a nonexistent event.
