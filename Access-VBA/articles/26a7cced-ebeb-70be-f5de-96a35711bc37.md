
# Close Method (ADO)

 **Last modified:** June 29, 2011

 _ **Applies to:** Access 2013 | Access 2016_



Closes an open object and any dependent objects.

## Syntax

 _object_. **Close**


## Remarks

Use the  **Close** method to close a[Connection](c16023aa-0321-2513-ee71-255d6ffba03d.md), a [Record](817aaf13-78d4-1134-aa94-997e92077c22.md), a [Recordset](0f963bf8-f066-dc8a-b754-f427de712df1.md), or a [Stream](d49b1514-e0b4-0aca-d5c2-8266f3f4fe65.md) object to free any associated system resources. Closing an object does not remove it from memory; you can change its property settings and open it again later. To completely eliminate an object from memory, set the object variable to _Nothing_ (in Visual Basic) after closing the object.

 **Connection**

Using the  **Close** method to close a **Connection** object also closes any active **Recordset** objects associated with the connection. A[Command](64f4ef03-f858-c004-b891-0c96d13a5e6e.md) object associated with the **Connection** object you are closing will persist, but it will no longer be associated with a **Connection** object; that is, its[ActiveConnection](5501b2d7-b62c-5fff-1edd-2b7efb3f8c4a.md) property will be set to **Nothing**. Also, the **Command** object's[Parameters](554387c3-3572-5391-3b24-c7d3443844cd.md) collection will be cleared of any provider-defined parameters.

You can later call the [Open](1adaa17d-dfe1-22e0-3415-720516d138f8.md) method to re-establish the connection to the same, or another, data source. While the **Connection** object is closed, calling any methods that require an open connection to the data source generates an error.

Closing a  **Connection** object while there are open **Recordset** objects on the connection rolls back any pending changes in all of the **Recordset** objects. Explicitly closing a **Connection** object (calling the **Close** method) while a transaction is in progress generates an error. If a **Connection** object falls out of scope while a transaction is in progress, ADO automatically rolls back the transaction.

 **Recordset, Record, Stream**

Using the  **Close** method to close a **Recordset**, **Record**, or **Stream** object releases the associated data and any exclusive access you may have had to the data through this particular object. You can later call the[Open](87ef19a4-28e1-dec7-ed33-4ae500b9c460.md) method to reopen the object with the same, or modified, attributes.

While a  **Recordset** object is closed, calling any methods that require a live cursor generates an error.

If an edit is in progress while in immediate update mode, calling the  **Close** method generates an error; instead, call the[Update](fc88cab6-c379-bb4f-530c-da08107924e0.md) or[CancelUpdate](2bd4d168-ba52-7786-5046-44febeda88e1.md) method first. If you close the **Recordset** object while in batch update mode, all changes since the last[UpdateBatch](69e72a65-b637-36fd-d09f-7f81050f71ad.md) call are lost.

If you use the [Clone](ca9b2b76-90bf-9a60-2611-3cb4977d5591.md) method to create copies of an open **Recordset** object, closing the original or a clone does not affect any of the other copies.

