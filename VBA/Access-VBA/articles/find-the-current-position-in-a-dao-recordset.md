---
title: Find the Current Position in a DAO Recordset
ms.prod: access
ms.assetid: 1f08caa7-b671-b844-59a0-f924a5220cf4
ms.date: 06/08/2017
---


# Find the Current Position in a DAO Recordset

In some situations, you need to determine how far through a  **[Recordset](http://msdn.microsoft.com/library/9774232C-E6DA-175B-FC7F-ED2AB7908FA0%28Office.15%29.aspx)** object you have moved the current record position, and perhaps indicate the current record position to a user. For example, you may want to indicate the current position on a dial, meter, or similar type of control. Two properties are available to indicate the current position: the **[AbsolutePosition](http://msdn.microsoft.com/library/C35C0C07-F789-524B-0A3D-DFD18FA6EEBC%28Office.15%29.aspx)** property and the **[PercentPosition](http://msdn.microsoft.com/library/AEBBDA44-ED72-7A6C-0CD5-28C8997D4D96%28Office.15%29.aspx)** property.

The  **AbsolutePosition** property value is the position of the current record relative to 0. However, do not think of this property as a record number; if the current record is undefined, the **AbsolutePosition** property returns 1. In addition, there is no assurance that a record will have the same absolute position if the **Recordset** object is recreated because the order of individual records within a **Recordset** object is not guaranteed unless it is created with an SQL statement that includes an ORDER BY clause.

The  **PercentPosition** property shows the current position expressed as a percentage of the total number of records indicated by the **[RecordCount](http://msdn.microsoft.com/library/AA1FED4F-CA51-918F-0A46-2B755B5F861A%28Office.15%29.aspx)** property. Because the **RecordCount** property does not reflect the total number of records in the **Recordset** object until the **Recordset** has been fully populated, the **PercentPosition** property reflects only the current record position as a percentage of the number of records that have been accessed since the **Recordset** was opened. To make sure that the **PercentPosition** property reflects the current record position relative to the entire **Recordset**, use the **[MoveLast](http://msdn.microsoft.com/library/FC0F7A33-1F55-9F5B-B00D-1B81F49B1C3E%28Office.15%29.aspx)** and **[MoveFirst](http://msdn.microsoft.com/library/338F7E86-6997-B80A-FC7A-A395D10B4A62%28Office.15%29.aspx)** methods immediately after opening the **Recordset**. This fully populates the **Recordset** object before you use the **PercentPosition** property. If you have a large result set, using the **MoveLast** method may take a long time for **Recordsets** that are not of type table.


 **Note**   The **PercentPosition** property is only an approximation and should not be used as a critical parameter. This property is best suited for driving an indicator that marks a user's progress while moving though a set of records. For example, you may want a control that indicates the percentage of records completed.

The following example opens a  **Recordset** object on a table called Employees. The procedure then moves through the Employees table and uses the **[SysCmd](application-syscmd-method-access.md)** method to display a progress bar showing the percentage of the table that has been processed. If the hire date of the employee is before Jan. 1, 1993, the text "Senior Staff" is appended to the Notes field.



```vb
Sub AddEmployeeNotes() 
 
Dim dbsNorthwind As DAO.Database 
Dim rstEmployees As DAO.Recordset 
Dim strMsg As String 
Dim intRet As Integer 
Dim intCount As Integer 
Dim strSQL As String 
Dim sngPercent As Single 
Dim varReturn As Variant 
Dim lngEmpID() As Long 
 
On Error GoTo ErrorHandler 
 
   Set dbsNorthwind = CurrentDb 
 
   strSQL = "SELECT * FROM Employees" 
   Set rstEmployees = dbsNorthwind.OpenRecordset(strSQL, dbOpenDynaset) 
 
   With rstEmployees 
      If .EOF Then            ' If no records, exit. 
         Exit Sub 
      Else 
         strMsg = "Processing Employees table..." 
         intRet = SysCmd(acSysCmdInitMeter, strMsg, 100) 
      End If 
 
      Do Until .EOF 
         If !HireDate < #1/1/93# Then 
            .Edit 
            !Notes = !Notes &; ";" &; "Senior Staff" 
            .Update 
         End If 
 
         If .PercentPosition <> 0 Then 
            intRet = SysCmd(acSysCmdUpdateMeter, .PercentPosition) 
         End If 
         .MoveNext 
      Loop 
   End With 
 
   intRet = SysCmd(acSysCmdRemoveMeter) 
 
   rstEmployees.Close 
   dbsNorthwind.Close 
 
   Set rstEmployees = Nothing 
   Set dbsNorthwind = Nothing 
 
Exit Sub 
 
ErrorHandler: 
   MsgBox "Error #: " &; Err.Number &; vbCrLf &; vbCrLf &; Err.Description 
   varReturn = SysCmd(acSysCmdSetStatus, " ") 
End Sub
```


