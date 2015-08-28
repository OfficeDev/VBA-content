
# Nothing <keyword>

 **Last modified:** July 28, 2015

The  **Nothing** [keyword](b8bdf64f-5920-1ae9-16d0-b26d09524a30.md) is used to disassociate an object [variable](b8bdf64f-5920-1ae9-16d0-b26d09524a30.md) from an actual object. Use the **Set** statement to assign **Nothing** to an object variable. For example:



```
Set MyObject = Nothing 

```

Several object variables can refer to the same actual object. When  **Nothing** is assigned to an object variable, that variable no longer refers to an actual object. When several object variables refer to the same object, memory and system resources associated with the object to which the variables refer are released only after all of them have been set to **Nothing**, either explicitly using  **Set**, or implicitly after the last object variable set to  **Nothing** goes out of [scope](b8bdf64f-5920-1ae9-16d0-b26d09524a30.md).
