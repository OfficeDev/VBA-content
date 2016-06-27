
# State Property (ADO)

 **Last modified:** June 29, 2011

 _ **Applies to:** Access 2013 | Access 2016_



Indicates for all applicable objects whether the state of the object is open or closed.
Indicates for all applicable objects executing an asynchronous method, whether the current state of the object is connecting, executing, or retrieving.

## Return Value

Returns a  **Long** value that can be an[ObjectStateEnum](129d589a-2955-3da9-e60a-7fbfdd6bfbdc.md) value. The default value is **adStateClosed**.


## Remarks

You can use the  **State** property to determine the current state of a given object at any time.

The object's  **State** property can have a combination of values. For example, if a statement is executing, this property will have a combined value of **adStateOpen** and **adStateExecuting**.

The  **State** property is read-only.

