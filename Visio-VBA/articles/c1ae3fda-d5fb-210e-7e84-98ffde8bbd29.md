
# Application.ShapeLinkDeleted Event (Visio)

 **Last modified:** July 28, 2015

 _**Applies to:** Visio 2013 Preview_

Occurs after the link between a shape and a data row is deleted.


 **Note**  This Visio object or member is available only to licensed users of Visio Professional 2013.


## Syntax

Private Sub  _expression__**ShapeLinkDeleted**( **_ByVal Shape As [IVSHAPE]_**,  **_ByVal DataRecordsetID As Long_**,  **_ByVal DataRowID As Long_**)

 _expression_An expression that returns a  **Application** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
|Shape|Required| **[IVSHAPE]**|The shape whose link to a data row was broken.|
|DataRecordsetID|Required| **Long**|The ID of the data recordset containing the data row that was linked to the shape.|
|DataRowID|Required| **Long**|The ID of the data row that was linked to the shape.|

## Remarks

The  **ShapeLinkDeleted** event is one of a group of events for which the **EventInfo** property of the **Application** object contains extra information.

When the  **ShapeLinkDeleted** event is fired, the **EventInfo** property returns the following string:

 `/DataRecordsetID = n /DataRowID = m`

where  _n_ and _m_ represent the IDs of the data recordset and data row, respectively, associated with the event.

If you're using Microsoft Visual Basic or Visual Basic for Applications (VBA), the syntax in this topic describes a common, efficient way to handle events.

If you want to create your own  **Event** objects, use the **Add** or **AddAdvise** method. To create an **Event** object that runs an add-on, use the **Add** method as it applies to the **EventList** collection. To create an **Event** object that receives notification, use the **AddAdvise** method. To find an event code for the event you want to create, see [Event codes](de8f5c7a-421d-ebcf-22b6-4310a202ef64.md).

