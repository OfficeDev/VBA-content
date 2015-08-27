
# DataRecordsetChangedEvent.DataColumnsDeleted Property (Visio)

 **Last modified:** July 28, 2015

 _**Applies to:** Visio 2013 Preview_

After data in a data recordset are refreshed, returns an array of names of data columns deleted from the data recordset as a result of the refresh operation. Read-only.


 **Note**  This Visio object or member is available only to licensed users of Visio Professional 2013.


## Syntax

 _expression_. **DataColumnsDeleted**

 _expression_An expression that returns a  **DataRecordsetChangedEvent** object.


### Return Value

String()


## Remarks

The columns returned by this property have already been deleted. As a result, you can no longer use Visio Automation properties or methods to retrieve the  **DataColumn** objects that represented the columns, nor any information about data formerly in these columns.

