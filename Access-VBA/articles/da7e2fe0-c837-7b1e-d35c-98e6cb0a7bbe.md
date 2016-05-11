
# DBEngine.Rollback Method (DAO)

 **Last modified:** June 29, 2011

 _ **Applies to:** Access 2013 | Access 2016_

Ends the current transaction and restores the databases in the  **Workspace** object to the state they were in when the current transaction began.


## Syntax

 _expression_. **Rollback**

 _expression_ A variable that represents a **DBEngine** object.


## Remarks

The transaction methods  **BeginTrans**, **CommitTrans**, and **Rollback** manage transaction processing during a session defined by a **Workspace** object. You use these methods with a **Workspace** object when you want to treat a series of changes made to the databases in a session as one unit.

Typically, you use transactions to maintain the integrity of your data when you must both update records in two or more tables and ensure changes are completed (committed) in all tables or none at all (rolled back). For example, if you transfer money from one account to another, you might subtract an amount from one and add the amount to another. If either update fails, the accounts no longer balance. Use the  **BeginTrans** method before updating the first record, and then, if any subsequent update fails, you can use the **Rollback** method to undo all of the updates. Use the **CommitTrans** method after you successfully update the last record.


 **Note**  Within one  **Workspace** object, transactions are always global to the **Workspace** and aren't limited to only one **Connection** or **Database** object. If you perform operations on more than one connection or database within a **Workspace** transaction, resolving the transaction (that is, using the **CommitTrans** or **Rollback** method) affects all operations on all connections and databases within that workspace.

After you use  **CommitTrans**, you can't undo changes made during that transaction unless the transaction is nested within another transaction that is itself rolled back. If you nest transactions, you must resolve the current transaction before you can resolve a transaction at a higher level of nesting.

If you want to have simultaneous transactions with overlapping, non-nested scopes, you can create additional  **Workspace** objects to contain the concurrent transactions.

If you close a  **Workspace** object without resolving any pending transactions, the transactions are automatically rolled back.

If you use the  **CommitTrans** or **Rollback** method without first using the **BeginTrans** method, an error occurs.

Some ISAM databases used in a Microsoft Access workspace may not support transactions, in which case the  **Transactions** property of the **Database** object or **Recordset** object is **False**. To make sure the database supports transactions, check the value of the **Transactions** property of the **Database** object before using the **BeginTrans** method. If you are using a **Recordset** object based on more than one database, check the **Transactions** property of the **Recordset** object. If a **Recordset** is based entirely on Microsoft Access database engine tables, you can always use transactions. **Recordset** objects based on tables created by other database products, however, may not support transactions. For example, you can't use transactions in a **Recordset** based on a Paradox table. In this case, the **Transactions** property is **False**. If the **Database** or **Recordset** doesn't support transactions, the methods are ignored and no error occurs.

You can't nest transactions if you are accessing ODBC data sources through the Microsoft Access database engine.

In ODBC workspaces, when you use  **CommitTrans** your cursor may no longer be valid. Use the **Requery** method to view the changes in the **Recordset**, or close and re-open the **Recordset**.


- You can often improve the performance of your application by breaking operations that require disk access into transaction blocks. This buffers your operations and may significantly reduce the number of times a disk is accessed.
    
- In a Microsoft Access workspace, transactions are logged in a file kept in the directory specified by the TEMP environment variable on the workstation. If the transaction log file exhausts the available storage on your TEMP drive, the database engine triggers a run-time error. At this point, if you use  **CommitTrans**, an indeterminate number of operations are committed, but the remaining uncompleted operations are lost, and the operation has to be restarted. Using a **Rollback** method releases the transaction log and rolls back all operations in the transaction.
    
- Closing a clone  **Recordset** within a pending transaction will cause an implicit **Rollback** operation.
    
