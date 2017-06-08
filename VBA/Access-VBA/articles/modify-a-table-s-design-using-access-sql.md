---
title: Modify a Table's Design Using Access SQL
ms.prod: access
ms.assetid: c05687af-ed43-56dc-a65a-e9c328be0f5b
ms.date: 06/08/2017
---


# Modify a Table's Design Using Access SQL

After you have created and populated a table, you may need to modify the table's design. To do so, use the  **[ALTER TABLE](http://msdn.microsoft.com/library/78E6C92C-E88C-E55F-6B89-435360C166A6%28Office.15%29.aspx)** statement. Be aware that altering an existing table's structure may cause you to lose some of the data. For example, changing a field's data type can result in data loss or rounding errors, depending on the data types you are using. It can also break other parts of your application that may refer to the changed field. You should always use extra caution before modifying the structure of an existing table.

With the  **ALTER TABLE** statement, you can add, remove, or change a column (or field), and you can add or remove a constraint. You can also declare a default value for a field; however, you can alter only one field at a time. Suppose that you have an invoicing database, and you want to add a field to the Customers table. To add a field with the **ALTER TABLE** statement, use the **ADD COLUMN** clause with the name of the field, its data type, and the size of the data type, if it is required.



```sql
ALTER TABLE tblCustomers 
   ADD COLUMN Address TEXT(30) 

```

To change the data type or size of a field, use the  **ALTER COLUMN** clause with the name of the field, the desired data type, and the desired size of the data type, if it is required.



```sql
ALTER TABLE tblCustomers 
   ALTER COLUMN Address TEXT(40) 

```

If you want to change the name of a field, you will have to remove the field and then recreate it. To remove a field, use the  **DROP COLUMN** clause with the field name only.



```sql
ALTER TABLE tblCustomers 
   DROP COLUMN Address 

```

Note that using this method will eliminate the existing data for the field. To preserve the existing data, you should change the field's name with the table design mode of the Access user interface, or write code to preserve the current data in a temporary table and append it back to the renamed table.
A default value is the value that is entered in a field any time a new record is added to a table and no value is specified for that particular column. To set a default value for a field, use the  **DEFAULT** keyword after declaring the field type in either an **ADD COLUMN** or **ALTER COLUMN** clause.



```sql
ALTER TABLE tblCustomers 
   ALTER COLUMN Address TEXT(40) DEFAULT Unknown 

```

Be aware that the default value is not enclosed in single quotation marks. If it were, the quotation marks would also be inserted into the record. The  **DEFAULT** keyword can also be used in a **[CREATE TABLE](http://msdn.microsoft.com/library/FC45D36E-6E43-C030-5016-CCA8BB1379FE%28Office.15%29.aspx)** statement.



```sql
CREATE TABLE tblCustomers ( 
   CustomerID INTEGER CONSTRAINT PK_tblCustomers 
      PRIMARY KEY,  
   [Last Name] TEXT(50) NOT NULL, 
   [First Name] TEXT(50) NOT NULL, 
   Phone TEXT(10), 
   Email TEXT(50), 
   Address TEXT(40) DEFAULT Unknown) 

```


 **Note**  The DEFAULT statement can be executed only through the Access OLE DB provider and ADO. It will return an error message if used through the Access SQL View user interface.


## Constraints

Constraints can be used to establish primary keys and referential integrity, and to restrict values that can be inserted into a field. In general, constraints can be used to preserve the integrity and consistency of the data in your database.

There are two types of constraints: a single-field or field-level constraint, and a multi-field or table-level constraint. Both kinds of constraints can be used in either the  **CREATE TABLE** or the **ALTER TABLE** statement.

A single-field constraint, also known as a column-level constraint, is declared with the field itself, after the field and data type have been declared. For this example, use the Customers table and create a single-field primary key on the CustomerID field. To add the constraint, use the  **CONSTRAINT** keyword with the name of the field.




```sql
ALTER TABLE tblCustomers 
   ALTER COLUMN CustomerID INTEGER 
   CONSTRAINT PK_tblCustomers PRIMARY KEY 

```

Be aware that the name of the constraint is given. You could use a shortcut for declaring the primary key that omits the  **CONSTRAINT** clause entirely.




```sql
ALTER TABLE tblCustomers 
   ALTER COLUMN CustomerID INTEGER PRIMARY KEY 

```

However, using the shortcut method will cause Access to randomly generate a name for the constraint, which will make it difficult to reference in code. It is a good idea always to name your constraints.

To drop a constraint, use the  **DROP CONSTRAINT** clause with the **ALTER TABLE** statement, and supply the name of the constraint.




```sql
ALTER TABLE tblCustomers 
   DROP CONSTRAINT PK_tblCustomers 

```

Constraints also can be used to restrict the allowable values for a field. You can restrict values to  **NOT NULL** or **UNIQUE**, or you can define a check constraint, which is a type of business rule that can be applied to a field. Imagine that you want to restrict (or constrain) the values of the first name and last name fields to be unique, meaning that there should never be a combination of first name and last name that is the same for any two records in the table. Because this is a multi-field constraint, it is declared at the table level, not the field level. Use the **ADD CONSTRAINT** clause and define a multi-field list.




```sql
ALTER TABLE tblCustomers 
   ADD CONSTRAINT CustomerID UNIQUE 
   ([Last Name], [First Name]) 

```

A check constraint is a powerful SQL feature that allows you to add data validation to a table by creating an expression that can refer to a single field, or multiple fields across one or more tables. Suppose that you want to make sure that the amounts entered in an invoice record are always greater than $0.00. To do so, use a check constraint by declaring the  **CHECK** keyword and your validation expression in the **ADD CONSTRAINT** clause of an **ALTER TABLE** statement.




```sql
ALTER TABLE tblInvoices 
   ADD CONSTRAINT CheckAmount 
   CHECK (Amount > 0) 

```

The expression used to define a check constraint also can refer to more than one field in the same table, or to fields in other tables, and can use any operations that are valid in Microsoft Access SQL, such as  **[SELECT](http://msdn.microsoft.com/library/A5C9DA94-5F9E-0FC0-767A-4117F38A5EF3%28Office.15%29.aspx)** statements, mathematical operators, and aggregate functions. The expression that defines the check constraint can be no more than 64 characters long.

Suppose that you want to check each customer's credit limit before he or she is added to the Customers table. Using an  **ALTER TABLE** statement with the **ADD COLUMN** and **CONSTRAINT** clauses, create a constraint that will look up the value in the CreditLimit table to verify the customer's credit limit. Use the following SQL statements to create the tblCreditLimit table, add the CustomerLimit field to the tblCustomers table, add the check constraint to the tblCustomers table, and test the check constraint.




```sql
CREATE TABLE tblCreditLimit ( 
   Limit DOUBLE) 
 
INSERT INTO tblCreditLimit 
   VALUES (100) 
 
ALTER TABLE tblCustomers 
   ADD COLUMN CustomerLimit DOUBLE 
 
ALTER TABLE tblCustomers 
   ADD CONSTRAINT LimitRule 
   CHECK (CustomerLimit <= (SELECT Limit 
      FROM tblCreditLimit)) 
 
UPDATE TABLE tblCustomers 
   SET CustomerLimit = 200 
   WHERE CustomerID = 1 

```

Be aware that when you execute the  **[UPDATE TABLE](http://msdn.microsoft.com/library/08F9C3D6-C020-ECF1-5748-43B93A76DFBB%28Office.15%29.aspx)** statement, you receive a message indicating that the update did not succeed because it violated the check constraint. If you update the CustomerLimit field to a value that is equal to or less than 100, the update will succeed.


## Cascading updates and deletions

Constraints also can be used to establish referential integrity between database tables. Having referential integrity means that the data is consistent and uncorrupted. For example, if you deleted a customer record but that customer's shipping record remained in the database, the data would be inconsistent because you now have an orphaned record in the shipping table. Referential integrity is established when you build a relationship between tables. In addition to establishing referential integrity, you can also ensure that the records in the referenced tables stay in sync by using cascading updates and deletions. For example, when the cascading updates and deletes are declared, if you delete the customer record, the customer's shipping record is deleted automatically.

To enable cascading updates and deletions, use the  **ON UPDATE CASCADE** and/or **ON DELETE CASCADE** keywords in the **CONSTRAINT** clause of an **ALTER TABLE** statement. Be aware that they must be applied to the foreign key.




```sql
ALTER TABLE tblShipping 
   ADD CONSTRAINT FK_tblShipping 
   FOREIGN KEY (CustomerID) REFERENCES 
      tblCustomers (CustomerID) 
   ON UPDATE CASCADE 
   ON DELETE CASCADE 

```


