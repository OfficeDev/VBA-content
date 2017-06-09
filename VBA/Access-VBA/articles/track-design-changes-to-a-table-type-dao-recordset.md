---
title: Track Design Changes to a Table-Type DAO Recordset
ms.prod: access
ms.assetid: 540bfd6d-14f1-e0a2-8a2e-e09cb1a31d52
ms.date: 06/08/2017
---


# Track Design Changes to a Table-Type DAO Recordset

You may need to determine when the underlying  **[TableDef](http://msdn.microsoft.com/library/715146B6-C62A-ABFF-28EE-E6BBE3C08ADF%28Office.15%29.aspx)** object of a table-type **[Recordset](http://msdn.microsoft.com/library/9774232C-E6DA-175B-FC7F-ED2AB7908FA0%28Office.15%29.aspx)** was created, or the last time it was modified. The **[DateCreated](http://msdn.microsoft.com/library/BD63AC73-2218-B62C-A785-DE08C4625DFF%28Office.15%29.aspx)** and **[LastUpdated](http://msdn.microsoft.com/library/091A8E10-01C0-20AF-7230-CD7103C243A1%28Office.15%29.aspx)** properties, respectively, give you this information. Both properties return the date stamp applied to the table by the machine on which the table resided at the time it was stamped. These properties are updated only when the table's design changes; they are not affected by changes to records in the table.

The following code example shows the  **DateCreated** and **LastUpdated** properties by adding a new **[Field](http://msdn.microsoft.com/library/1DBD535E-48AD-A5C8-A1B2-6776C1E3E19D%28Office.15%29.aspx)** to an existing **TableDef** and by creating a new **TableDef**. The **DateOutput** function is required for this procedure to run.



```vb
Sub DateCreatedX() 
 
   Dim dbsNorthwind As Database 
   Dim tdfEmployees As TableDef 
   Dim tdfNewTable As TableDef 
 
   Set dbsNorthwind = OpenDatabase("Northwind.mdb") 
 
   With dbsNorthwind 
      Set tdfEmployees = .TableDefs!Employees 
 
      With tdfEmployees 
         ' Print current information about the Employees  
         ' table. 
         DateOutput "Current properties", tdfEmployees 
 
         ' Create and append a field to the Employees table. 
         .Fields.Append .CreateField("NewField", dbDate) 
 
         ' Print new information about the Employees  
         ' table. 
         DateOutput "After creating a new field", _ 
            tdfEmployees 
 
         ' Delete new Field because this is a demonstration. 
         .Fields.Delete "NewField" 
      End With 
 
      ' Create and append a new TableDef object to the  
      ' Northwind database. 
      Set tdfNewTable = .CreateTableDef("NewTableDef") 
      With tdfNewTable 
         .Fields.Append .CreateField("NewField", dbDate) 
      End With 
      .TableDefs.Append tdfNewTable 
 
      ' Print information about the new TableDef object. 
      DateOutput "After creating a new table", tdfNewTable 
 
      ' Delete new TableDef object because this is a  
      ' demonstration. 
      .TableDefs.Delete tdfNewTable.Name 
      .Close 
   End With 
 
End Sub 
 
Function DateOutput(strTemp As String, _ 
   tdfTemp As TableDef) 
 
   ' Print DateCreated and LastUpdated information about  
   ' specified TableDef object. 
   Debug.Print strTemp 
   Debug.Print "  TableDef: " &; tdfTemp.Name 
   Debug.Print "    DateCreated = " &; _ 
      tdfTemp.DateCreated 
   Debug.Print "    LastUpdated = " &; _ 
      tdfTemp.LastUpdated 
   Debug.Print 
 
End Function
```


