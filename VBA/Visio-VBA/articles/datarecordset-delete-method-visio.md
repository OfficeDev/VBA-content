---
title: DataRecordset.Delete Method (Visio)
keywords: vis_sdr.chm16416165
f1_keywords:
- vis_sdr.chm16416165
ms.prod: visio
api_name:
- Visio.DataRecordset.Delete
ms.assetid: 9f3fa9b0-2ca9-cf28-fa27-18eef4be179d
ms.date: 06/08/2017
---


# DataRecordset.Delete Method (Visio)

Deletes the  **[DataRecordset](datarecordset-object-visio.md)** object from the **[DataRecordsets](datarecordsets-object-visio.md)** collection of the document. .


 **Note**  This Visio object or member is available only to licensed users of Visio Professional 2013.


## Syntax

 _expression_ . **Delete**

 _expression_ A variable that represents a **DataRecordset** object.


### Return Value

Nothing


## Remarks

If the  **DataRecordset** object to be deleted is associated with a **[DataConnection](dataconnection-object-visio.md)** object, and if that **DataConnection** object is not associated with any other **DataRecordset** objects, Microsoft Visio also deletes the **DataConnection** object.

Note that deleting a  **DataRecordset** object does not delete the shapes that had been linked to data in that data recordset, nor does delete any existing shape data in those shapes that was created when the shapes were linked to data.


## Example

This Microsoft Visual Basic for Applications (VBA) macro shows how to use the  **Delete** method to delete a **DataRecordset** from the **DataRecordsets** collection of the current document. It gets the count of all data recordsets associated with the current document and deletes the one most recently added.


```vb
Public Sub Delete_Example() 
 
    Dim vsoDataRecordset As Visio.DataRecordset 
    Dim intCount As Integer 
 
    intCount = ThisDocument.DataRecordsets.Count 
    Set vsoDataRecordset = ThisDocument.DataRecordsets(intCount) 
    vsoDataRecordset.Delete 
 
End Sub
```


