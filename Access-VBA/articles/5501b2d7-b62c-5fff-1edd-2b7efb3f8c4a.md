
# ActiveConnection Property (ADO)

 **Last modified:** June 29, 2011

 _ **Applies to:** Access 2013 | Access 2016_



Indicates to which [Connection](c16023aa-0321-2513-ee71-255d6ffba03d.md) object the specified[Command](64f4ef03-f858-c004-b891-0c96d13a5e6e.md), [Recordset](0f963bf8-f066-dc8a-b754-f427de712df1.md), or [Record](817aaf13-78d4-1134-aa94-997e92077c22.md) object currently belongs.

## Settings and Return Values

Sets or returns a  **String** value that contains a definition for a connection if the connection is closed, or a **Variant** containing the current **Connection** object if the connection is open. Default is a null object reference. See the[ConnectionString](c67a7daf-258f-d99d-6475-a4aa98d1e99d.md) property.


## Remarks

Use the  **ActiveConnection** property to determine the **Connection** object over which the specified **Command** object will execute or the specified **Recordset** will be opened.

 **Command**

For  **Command** objects, the **ActiveConnection** property is read/write.

If you attempt to call the [Execute](http://msdn.microsoft.com/library/01812c8c-403e-4428-23f6-86bda747bd0e%28Office.15%29.aspx) method on a **Command** object before setting this property to an open **Connection** object or valid connection string, an error occurs.

 **Microsoft Visual Basic** Setting the **ActiveConnection** property to _Nothing_ disassociates the **Command** object from the current **Connection** and causes the provider to release any associated resources on the data source. You can then associate the **Command** object with the same or another **Connection** object. Some providers allow you to change the property setting from one **Connection** to another, without having to first set the property to _Nothing_.

If the [Parameters](554387c3-3572-5391-3b24-c7d3443844cd.md) collection of the **Command** object contains parameters supplied by the provider, the collection is cleared if you set the **ActiveConnection** property to _Nothing_ or to another **Connection** object. If you manually create[Parameter](7577598e-3d0c-30c6-5f24-1cfe98791798.md) objects and use them to fill the **Parameters** collection of the **Command** object, setting the **ActiveConnection** property to _Nothing_ or to another **Connection** object leaves the **Parameters** collection intact.

Closing the  **Connection** object with which a **Command** object is associated sets the **ActiveConnection** property to _Nothing_. Setting this property to a closed **Connection** object generates an error.

 **Recordset**

For open  **Recordset** objects or for **Recordset** objects whose[Source](523ea81e-d011-8d87-436e-084b6eba0908.md) property is set to a valid **Command** object, the **ActiveConnection** property is read-only. Otherwise, it is read/write.

You can set this property to a valid  **Connection** object or to a valid connection string. In this case, the provider creates a new **Connection** object using this definition and opens the connection. Additionally, the provider may set this property to the new **Connection** object to give you a way to access the **Connection** object for extended error information or to execute other commands.

If you use the  _ActiveConnection_ argument of the[Open](87ef19a4-28e1-dec7-ed33-4ae500b9c460.md) method to open a **Recordset** object, the **ActiveConnection** property will inherit the value of the argument.

If you set the  **Source** property of the **Recordset** object to a valid **Command** object variable, the **ActiveConnection** property of the **Recordset** inherits the setting of the **Command** object's **ActiveConnection** property.

 **Remote Data Service Usage** When used on a client-side Recordset object, this property can be set only to a connection string or (in Microsoft Visual Basic or Visual Basic, Scripting Edition) to _Nothing_.

 **Record**

This property is read/write when the  **Record** object is closed, and may contain a connection string or reference to an open **Connection** object. This property is read-only when the **Record** object is open, and contains a reference to an open **Connection** object.

A  **Connection** object is created implicitly when the **Record** object is opened from a URL. Open the **Record** with an existing, open **Connection** object by assigning the **Connection** object to this property, or using the **Connection** object as a parameter in the[Open](ba71c5c7-326e-d3b6-0e74-e8343ee6896f.md) method call. If the **Record** is opened from an existing **Record** or[Recordset](0f963bf8-f066-dc8a-b754-f427de712df1.md), then it is automatically associated with that  **Record** or **Recordset** object's **Connection** object.


 **Note**  URLs using the http scheme will automatically invoke the [Microsoft OLE DB Provider for Internet Publishing](5d1e8db5-dabb-0914-e11e-e2eac72bfa77.md). For more information, see [Absolute and Relative URLs](79a1f793-7154-1c13-7dfe-a1b8cd64e1ea.md).

