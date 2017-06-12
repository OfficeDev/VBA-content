---
title: Read From and Write To a Field in a DAO Recordset
ms.prod: access
ms.assetid: 4fe0c334-9c44-773c-7aed-182b042213a7
ms.date: 06/08/2017
---


# Read From and Write To a Field in a DAO Recordset

When you read or write data to a field, you are actually reading or setting the DAO  **[Value](http://msdn.microsoft.com/library/6C0F9A8D-F51A-B8CF-8830-F8D960A1D08C%28Office.15%29.aspx)** property of a **[Field](http://msdn.microsoft.com/library/47282CE2-9B49-CCF9-AD37-C4BB25CFD037%28Office.15%29.aspx)** object. The DAO **Value** property is the default property of a **Field** object. Therefore, you can set the DAO **Value** property of the LastName field in the rstEmployees **[Recordset](http://msdn.microsoft.com/library/9774232C-E6DA-175B-FC7F-ED2AB7908FA0%28Office.15%29.aspx)** in any of the following ways.


```
rstEmployees!LastName.Value = strName 
rstEmployees!LastName = strName 
rstEmployees![LastName] = strName 

```


The tables underlying a  **Recordset** object may not permit you to modify data, even though the **Recordset** is of type dynaset or table, which are usually updatable. Check the **[Updatable](http://msdn.microsoft.com/library/2D4BDCEF-1B10-B542-CE0F-6172C271131B%28Office.15%29.aspx)** property of the **Recordset** to determine whether its data can be changed. If the property is **True**, the **Recordset** object can be updated.

Individual fields within an updatable  **Recordset** object may not be updatable, and trying to write to these fields generates a run-time error. To determine whether a given field is updatable, check the **[DataUpdatable](http://msdn.microsoft.com/library/08CA57B6-2D7C-36B4-7D51-B76AC5467163%28Office.15%29.aspx)** property of the corresponding **Field** object in the **[Fields](http://msdn.microsoft.com/library/4BE3BA07-20C1-D958-C1B8-7DD8B4731F60%28Office.15%29.aspx)** collection of the **Recordset**. The following example returns **True** if all fields in the dynaset created by strQuery are updatable and returns **False** otherwise.



```vb
Function RecordsetUpdatable(strSQL As String) As Boolean 
 
Dim dbsNorthwind As DAO.Database 
Dim rstDynaset As DAO.Recordset 
Dim intPosition As Integer 
 
On Error GoTo ErrorHandler 
 
   ' Initialize the function's return value to True. 
   RecordsetUpdatable = True 
 
   Set dbsNorthwind = CurrentDb 
   Set rstDynaset = dbsNorthwind.OpenRecordset(strSQL, dbOpenDynaset) 
 
   ' If the entire dynaset isn't updatable, return False. 
   If rstDynaset.Updatable = False Then 
      RecordsetUpdatable = False 
   Else 
      ' If the dynaset is updatable, check if all fields in the 
      ' dynaset are updatable. If one of the fields isn't updatable, 
      ' return False. 
      For intPosition = 0 To rstDynaset.Fields.Count - 1 
         If rstDynaset.Fields(intPosition).DataUpdatable = False Then 
            RecordsetUpdatable = False 
            Exit For 
         End If 
      Next intPosition 
   End If 
 
   rstDynaset.Close 
   dbsNorthwind.Close 
 
   Set rstDynaset = Nothing 
   Set dbsNorthwind = Nothing 
 
Exit Sub 
 
ErrorHandler: 
   MsgBox "Error #: " &; Err.Number &; vbCrLf &; vbCrLf &; Err.Description 
End Function
```

Any single field can impose a number of criteria on data in that field when records are added or updated. These criteria are defined by a handful of properties. The DAO  **[AllowZeroLength](http://msdn.microsoft.com/library/5103A905-9258-E088-0210-857372F41C3C%28Office.15%29.aspx)** property on a Text or Memo field indicates whether or not the field will accept a zero-length string (""). The DAO **[Required](http://msdn.microsoft.com/library/2F1DBDEB-A37A-59B2-FDC2-F16C7AE1A575%28Office.15%29.aspx)** property indicates whether or not some value must be entered in the field, or if it instead can accept a **Null** value. For a **Field** object on a **Recordset**, these properties are read-only; their state is determined by the underlying table.
Validation is the process of determining whether data entered into a field's DAO  **Value** property is within an acceptable range. A **Field** object on a **Recordset** may have the DAO **[ValidationRule](http://msdn.microsoft.com/library/B07E644D-54D3-7199-6F99-178774E54398%28Office.15%29.aspx)** and **[ValidationText](http://msdn.microsoft.com/library/6D9EC790-A9D2-84D7-CCBA-57D738491E36%28Office.15%29.aspx)** properties set. The DAO **ValidationRule** property is simply a criteria expression, similar to the criteria of an SQL WHERE clause, without the WHERE keyword. The DAO **ValidationText** property is a string that Access displays in an error message if you try to enter data in the field that is outside the limits of the DAO **ValidationRule** property. If you are using DAO in your code, then you can use the DAO **ValidationText** for a message that you want to display to the user.

 **Note**  The DAO  **ValidationRule** and **ValidationText** properties also exist at the **Recordset** level. These are read-only properties, reflecting the table-level validation scheme established on the table from which the current record is retrieved.

A  **Field** object on a **Recordset** also features the **[ValidateOnSet](http://msdn.microsoft.com/library/00245A8A-A78F-B0A8-3EB3-11DD27873984%28Office.15%29.aspx)** property. When the **ValidateOnSet** property is set to **True**, Access checks validation as soon as the field's DAO **Value** property is set. When it is set to **False** (the default), Access checks validation only when the completed record is updated. For example, if you are adding data to a record that contains a large Memo or OLE Object field and that has the DAO **ValidationRule** property set, you should determine whether the new data violates the validation rule before trying to write the data. To do so, set the **ValidateOnSet** property to **True**. If you wait to check validation until the entire record is written to disk, you may waste time trying to write an invalid record to disk.

