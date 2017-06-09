---
title: Use Transactions in a DAO Recordset
ms.prod: access
ms.assetid: 7d565770-37b6-5650-c998-9ff3b30d54cb
ms.date: 06/08/2017
---


# Use Transactions in a DAO Recordset

A transaction is a set of operations bundled together and handled as a single unit of work. The work in a transaction must be completed as a whole; if any part of the transaction fails, the entire transaction fails. Transactions offer the developer the ability to enforce data integrity. With multiple database operations bundled into a single unit that must succeed or fail as a whole, the database cannot reach an inconsistent state. Transactions are common to most database management systems.

The most common example of transaction processing involves a bank's automated teller machine (ATM). The processes of dispensing cash and then debiting the user's account are considered a logical unit of work and are wrapped in a transaction: The cash is not dispensed unless the system is also able to debit the account. By using a transaction, the entire operation either succeeds or fails. This maintains the consistent state of the ATM database.

You should consider using transactions if you want to make sure that each operation in a group of operations is successful before all operations are committed. Keep in mind that all transactions are invisible to other transactions. That is, no transaction can see another transaction's updates to the database until the transaction is committed.


 **Note**  The behavior of transactions with Access databases differs from the behavior of ODBC data sources, such as SQL Server. For example, if a database is connected to a file server, and the file server stops before a transaction has had time to commit its changes, then your database could be left in an inconsistent state. If you require true transaction support with respect to durability, you should investigate using a client/server architecture.

The Access databse engine supports transactions through the DAO  **[BeginTrans](http://msdn.microsoft.com/library/FDDE662F-2472-3AF6-67D6-C8CA7FB1DCA7%28Office.15%29.aspx)**, **[CommitTrans](http://msdn.microsoft.com/library/0C9D345F-13FF-7FE6-789D-FBDB43FA54B8%28Office.15%29.aspx)**, and **[Rollback](http://msdn.microsoft.com/library/DA7E2FE0-C837-7B1E-D35C-98E6CB0A7BBE%28Office.15%29.aspx)** methods of the **[Workspace](http://msdn.microsoft.com/library/BF3AB863-5E9A-4842-1F82-2CCF958D9779%28Office.15%29.aspx)** object.
The following code example changes the job title of all sales representatives in the Employees table. After the  **BeginTrans** method starts a transaction that isolates all of the changes made to the Employees table, the **CommitTrans** method saves the changes. Be aware that you can use the **Rollback** method to undo changes that you saved with the **[Update](http://msdn.microsoft.com/library/AAD4171A-DA95-ED72-86B3-714615EA0AC8%28Office.15%29.aspx)** method.



```vb
Sub ChangeTitle() 
 
Dim wrkCurrent As DAO.Workspace 
Dim dbsNorthwind As DAO.Database 
Dim rstEmployee As DAO.Recordset 
 
On Error GoTo ErrorHandler 
 
   Set wrkCurrent = DBEngine.Workspaces(0) 
   Set dbsNorthwind = CurrentDB 
   Set rstEmployee = dbsNorthwind.OpenRecordset("Employees") 
 
   wrkCurrent.BeginTrans 
   Do Until rstEmployee.EOF 
      If rstEmployee!Title = "Sales Representative" Then 
         rstEmployee.Edit 
         rstEmloyee!Title = "Sales Associate" 
         rstEmployee.Update 
      End If 
      rstEmployee.MoveNext 
   Loop 
 
   If MsgBox("Save all changes?", vbQuestion + vbYesNo) = vbYes Then 
      wrkCurrent.CommitTrans 
   Else 
      wrkCurrent.Rollback 
   End If 
 
   rstEmployee.Close 
   dbsNorthwind.Close 
   wrkCurrent.Close 
 
   Set rstEmployee = nothing 
   Set dbsNorthwind = Nothing 
   Set wrkCurrent = Nothing 
 
   Exit Sub 
 
ErrorHandler: 
   MsgBox "Error #: " &; Err.Number &; vbCrLf &; vbCrLf &; Err.Description 
End Sub
```

When you use transactions, all databases and  **[Recordset](http://msdn.microsoft.com/library/9774232C-E6DA-175B-FC7F-ED2AB7908FA0%28Office.15%29.aspx)** objects in the specified **Workspace** object are affected; transactions are global to the workspace, not to a specific database or **Recordset**. If you perform operations on more than one database or within a workspace transaction, the **Commit** and **Rollback** methods affect all the objects changed within that workspace during the transaction.
You can also use the  **[BeginTrans](http://msdn.microsoft.com/library/FDDE662F-2472-3AF6-67D6-C8CA7FB1DCA7%28Office.15%29.aspx)**, **[CommitTrans](http://msdn.microsoft.com/library/0C9D345F-13FF-7FE6-789D-FBDB43FA54B8%28Office.15%29.aspx)**, and **[Rollback](http://msdn.microsoft.com/library/DA7E2FE0-C837-7B1E-D35C-98E6CB0A7BBE%28Office.15%29.aspx)** methods with the **[DBEngine](http://msdn.microsoft.com/library/CEAEB505-615E-37BA-4633-27240EF8C5DE%28Office.15%29.aspx)** object. In this case, the transaction is applied to the default workspace, which is `DBEngine.Workspaces(0)`.

