---
title: Make Bulk Changes to a DAO Recordset
ms.prod: access
ms.assetid: b66c857a-42ed-15c9-e01d-99c451651f3b
ms.date: 06/08/2017
---


# Make Bulk Changes to a DAO Recordset

After you have created a table-type or dynaset-type  **[Recordset](http://msdn.microsoft.com/library/9774232C-E6DA-175B-FC7F-ED2AB7908FA0%28Office.15%29.aspx)** object, you can change, delete, or add new records. You cannot change, delete, or add records to a snapshot-type or forward-only-type **Recordset** object.

Many of the changes you may otherwise perform in a loop can be done more efficiently with an update or delete query using SQL. The following example creates a  **[QueryDef](http://msdn.microsoft.com/library/0B3D901C-345D-42A2-F5F1-FB09CC562E27%28Office.15%29.aspx)** object to update the Employees table and then runs the query.



```vb
Dim dbsNorthwind As DAO.Database 
Dim qdfChangeTitles As DAO.QueryDef 
 
   Set dbsNorthwind = CurrentDb 
   Set qdfChangeTitles = dbsNorthwind.CreateQueryDef("") 
 
   qdfChangeTitles.SQL = "UPDATE Employees SET Title = " &; _ 
                         "'Account Executive' WHERE Title = " &; _ 
                         "'Sales Representative'" 
 
   ' Invoke query. 
   qdfChangeTitles.Execute dbFailOnError 

```

You can replace the entire SQL string in this example with a stored parameter query, in which case the procedure would prompt the user for parameter values. The following example shows how the previous example may be rewritten as a stored parameter query.



```vb
Dim dbsNorthwind As DAO.Database 
Dim qdfChangeTitles As DAO.QueryDef 
Dim strSQLUpdate As String 
Dim strOld As String 
Dim strNew As String 
 
   Set dbsNorthwind = CurrentDb 
 
   strSQLUpdate = "PARAMETERS [Old Title] Text, [New Title] Text; " &; _ 
                  "UPDATE Employees SET Title = [New Title] WHERE " &; _ 
                  "Title = [Old Title]" 
 
   ' Create the unstored QueryDef object. 
   Set qdfChangeTitles = dbsNorthwind.CreateQueryDef("", strSQLUpdate) 
 
   ' Prompt for old title. 
   strOld = InputBox("Enter old job title:") 
 
   ' Prompt for new title. 
   strNew = InputBox("Enter new job title:") 
 
   ' Set parameters. 
   qdfChangeTitles.Parameters("Old Title") = strOld 
   qdfChangeTitles.Parameters("New Title") = strNew 
 
   ' Invoke query. 
   qdfChangeTitles.Execute 

```


 **Note**  A delete query is much more efficient than code that loops through a  **Recordset**, modifying or deleting one record at a time.


