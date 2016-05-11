
# ConnectionTimeout Property (ADO)

 **Last modified:** June 29, 2011

 _ **Applies to:** Access 2013 | Access 2016_



Indicates how long to wait while establishing a connection before terminating the attempt and generating an error.

## Settings and Return Values

Sets or returns a  **Long** value that indicates, in seconds, how long to wait for the connection to open. Default is 15.


## Remarks

Use the  **ConnectionTimeout** property on a[Connection](c16023aa-0321-2513-ee71-255d6ffba03d.md) object if delays from network traffic or heavy server use make it necessary to abandon a connection attempt. If the time from the **ConnectionTimeout** property setting elapses prior to the opening of the connection, an error occurs and ADO cancels the attempt. If you set the property to zero, ADO will wait indefinitely until the connection is opened. Make sure the provider to which you are writing code supports the **ConnectionTimeout** functionality.

The  **ConnectionTimeout** property is read/write when the connection is closed and read-only when it is open.

