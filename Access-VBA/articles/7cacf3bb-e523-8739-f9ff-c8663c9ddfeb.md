
# Transaction Processing

 **Last modified:** June 29, 2011

 _ **Applies to:** Access 2013 | Access 2016_

ADO provides the following methods for controlling transactions:  **BeginTrans**, **CommitTrans**, and **RollbackTrans**. Use these methods with a **Connection** object when you want to save or cancel a series of changes made to the source data as a single unit. For example, to transfer money between accounts, you subtract an amount from one and add the same amount to the other. If either update fails, the accounts no longer balance. Making these changes within an open transaction ensures that either all or none of the changes go through.


 **Note**  Not all providers support transactions. Verify that the provider-defined property " **Transaction DDL** " appears in the **Connection** object's[Properties](4d662790-1252-c930-e6f9-edf6a38636af.md) collection, indicating that the provider supports transactions. If the provider does not support transactions, calling one of these methods will return an error.

After you call the  **BeginTrans** method, the provider will no longer instantaneously commit changes you make until you call **CommitTrans** or **RollbackTrans** to end the transaction.
Calling the  **CommitTrans** method saves changes made within an open transaction on the connection and ends the transaction. Calling the **RollbackTrans** method reverses any changes made within an open transaction and ends the transaction. Calling either method when there is no open transaction generates an error.
Depending on the  **Connection** object's[Attributes](4cc1f036-606e-7d4b-d270-af374e9d99fa.md) property, calling either the **CommitTrans** or **RollbackTrans** method may automatically start a new transaction. If the **Attributes** property is set to **adXactCommitRetaining**, the provider automatically starts a new transaction after a **CommitTrans** call. If the **Attributes** property is set to **adXactAbortRetaining**, the provider automatically starts a new transaction after a **RollbackTrans** call.

## Transaction Isolation Level

Use the  **IsolationLevel** property to set the isolation level of a transaction on a **Connection** object. The setting does not take effect until the next time you call the[BeginTrans](9a0415f0-9424-8d1c-4779-92e932292d46.md) method. If the level of isolation you request is unavailable, the provider may return the next greater level of isolation. Refer to the **IsolationLevel** property in the ADO Programmer's Reference for more details on valid values.


## Nested Transactions

For providers that support nested transactions, calling the  **BeginTrans** method within an open transaction starts a new, nested transaction. The return value indicates the level of nesting: a return value of "1" indicates you have opened a top-level transaction (that is, the transaction is not nested within another transaction), "2" indicates that you have opened a second-level transaction (a transaction nested within a top-level transaction), and so forth. Calling **CommitTrans** or **RollbackTrans** affects only the most recently opened transaction; you must close or roll back the current transaction before you can resolve any higher-level transactions.

