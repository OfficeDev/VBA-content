
# DeleteRecord Method (ADO)

 **Last modified:** June 29, 2011

 _ **Applies to:** Access 2013 | Access 2016_

 **In this article**
[Syntax](#sectionSection1)
[Parameters](#sectionSection2)
[Remarks](#sectionSection3)



Deletes a the entity represented by a [Record](817aaf13-78d4-1134-aa94-997e92077c22.md).

## Syntax
<a name="sectionSection1"> </a>

 _Record_. **DeleteRecord** _Source_, _Async_


## Parameters
<a name="sectionSection2"> </a>


-  _Source_
    
- Optional. A  **String** value that contains a URL identifying the entity (for example, the file or directory) to be deleted. If _Source_ is omitted or specifies an empty string, the entity represented by the current[Record](817aaf13-78d4-1134-aa94-997e92077c22.md) is deleted. If the Record is a collection record ([RecordType](a42001a6-7312-162d-dd71-c82f8c9d527f.md) of **adCollectionRecord**, such as a directory) all children (for example, subdirectories) will also be deleted.
    
-  _Async_
    
- Optional. A  **Boolean** value that, when **True**, specifies that the delete operation is asynchronous.
    

## Remarks
<a name="sectionSection3"> </a>

Operations on the object represented by this  **Record** may fail after this method completes. After calling **DeleteRecord**, the **Record** should be closed because the behavior of the **Record** may become unpredictable depending upon when the provider updates the **Record** with the data source.

If this  **Record** was obtained from a[Recordset](0f963bf8-f066-dc8a-b754-f427de712df1.md), then the results of this operation will not be reflected immediately in the  **Recordset**. Refresh the **Recordset** by closing and re-opening it, or by executing the **Recordset**[Requery](1062d907-979f-020a-b2ed-94e11c0e7d08.md), or [Update](fc88cab6-c379-bb4f-530c-da08107924e0.md) and[Resync](f594a200-56e6-fcf5-9b0a-900c56377f24.md) methods.


 **Note**  URLs using the http scheme will automatically invoke the [Microsoft OLE DB Provider for Internet Publishing](5d1e8db5-dabb-0914-e11e-e2eac72bfa77.md). For more information, see [Absolute and Relative URLs](79a1f793-7154-1c13-7dfe-a1b8cd64e1ea.md).

