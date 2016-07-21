
# CommandText Property (ADO)

 **Last modified:** June 29, 2011

 _ **Applies to:** Access 2013 | Access 2016_



Indicates the text of a command to be issued against a provider.

## Settings and Return Values

Sets or returns a  **String** value that contains a provider command, such as an SQL statement, a table name, a relative URL, or a stored procedure call. Default is "" (zero-length string).


## Remarks

Use the  **CommandText** property to set or return the text of a command represented by a[Command](64f4ef03-f858-c004-b891-0c96d13a5e6e.md) object. Usually this will be an SQL statement, but can also be any other type of command statement recognized by the provider, such as a stored procedure call. An SQL statement must be of the particular dialect or version supported by the provider's query processor.

If the [Prepared](33becda2-faab-5000-8904-6ffd8c5805f2.md) property of the **Command** object is set to **True** and the **Command** object is bound to an open connection when you set the **CommandText** property, ADO prepares the query (that is, a compiled form of the query that is stored by the provider) when you call the[Execute](http://msdn.microsoft.com/library/01812c8c-403e-4428-23f6-86bda747bd0e%28Office.15%29.aspx) or **Open** methods.

Depending on the [CommandType](c8d4fc1c-502b-11f3-af9d-605a03b6f056.md) property setting, ADO may alter the **CommandText** property. You can read the **CommandText** property at any time to see the actual command text that ADO will use during execution.

Use the  **CommandText** property to set or return a relative URL that specifies a resource, such as a file or directory. The resource is relative to a location specified explicitly by an absolute URL, or implicitly by an open[Connection](c16023aa-0321-2513-ee71-255d6ffba03d.md) object.


 **Note**  URLs using the http scheme will automatically invoke the [Microsoft OLE DB Provider for Internet Publishing](5d1e8db5-dabb-0914-e11e-e2eac72bfa77.md). For more information, see [Absolute and Relative URLs](79a1f793-7154-1c13-7dfe-a1b8cd64e1ea.md).

