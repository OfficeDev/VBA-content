---
title: Manipulate Multivalued Fields With DAO
ms.prod: access
ms.assetid: a3c02fcd-ad48-c3fb-afa1-aabb43fc5bbf
ms.date: 06/08/2017
---


# Manipulate Multivalued Fields With DAO

Multivalued fields are represented as  **[Recordset](http://msdn.microsoft.com/library/9774232C-E6DA-175B-FC7F-ED2AB7908FA0%28Office.15%29.aspx)** objects in DAO. The recordset for a field is a child of the recordset for the table that contains the multivalued field. To instantiate the child recordset, use the **Value** property of the multivalued field as follows.


```vb
Set childRs = rs.<multi-valued field>.Value
```


The following code example shows how to instantiate the child recordset of the AssignedTo field of the Tasks table.




```vb
Set rs  = db.OpenRecordSet("Tasks") 
Set childRs = rs.AssignedTo.Value 

```

The child recordset has the same functionality as any DAO  **Recordset** object.
The following code example shows how to iterate through a parent recordset and its child recordset. The example prints the tasks in the Tasks table along with the people assigned to the tasks to the Immediate window.



```vb
Sub BrowseMultiValueField() 
   Dim db As Database 
   Dim rs As Recordset 
   Dim childRS As Recordset 
     
   Set db = CurrentDb() 
     
   ' Open a Recordset for the Tasks table. 
   Set rs = db.OpenRecordset("Tasks") 
   rs.MoveFirst 
     
   Do Until rs.EOF 
      ' Print the name of the task to the Immediate window. 
      Debug.Print rs!TaskName.Value 
         
      ' Open a Recordset for the multivalued field. 
      Set childRS = rs!AssignedTo.Value 
 
         ' Exit the loop if the multivalued field contains no records. 
         Do Until childRS.EOF 
             childRS.MoveFirst 
                     
             ' Loop through the records in the child recordset. 
             Do Until childRS.EOF 
                 ' Print the owner(s) of the task to the Immediate  
                 ' window. 
                 Debug.Print Chr(0), childRS!Value.Value 
                 childRS.MoveNext 
             Loop 
         Loop 
      rs.MoveNext 
   Loop 
End Sub
```


