---
title: Define Relationships Between Tables Using Access SQL
ms.prod: access
ms.assetid: 24159c8a-c4ba-79a9-2490-007a82163f55
ms.date: 06/08/2017
---


# Define Relationships Between Tables Using Access SQL

Relationships are the established associations between two or more tables. Relationships are based on common fields from more than one table, often involving primary and foreign keys.

A primary key is the field (or fields) that is used to uniquely identify each record in a table. There are three requirements for a primary key: It cannot be null, it must be unique, and there can be only one defined per table. You can define a primary key either by creating a primary key index after the table is created, or by using the  **CONSTRAINT** clause in the table declaration, as shown in the examples later in this section. A constraint limits (or constrains) the values that are entered in a field.

A foreign key is a field (or fields) in one table that references the primary key in another table. The data in the fields from both tables is exactly the same, and the table with the primary key record (the primary table) must have existing records before the table with the foreign key record (the foreign table) has the matching or related records. Like primary keys, you can define foreign keys in the table declaration by using the  **CONSTRAINT** clause.

There are essentially three types of relationships: 

-  **One-to-one** For every record in the primary table, there is one and only one record in the foreign table.
    
-  **One-to-many** For every record in the primary table, there are one or more related records in the foreign table.
    
-  **Many-to-many** For every record in the primary table, there are many related records in the foreign table, and for every record in the foreign table, there are many related records in the primary table.
    
For example, suppose you want to add an invoices table to an invoicing database. Every customer in your customers table can have many invoices in the invoices table—this is a classic one-to-many scenario. You can take the primary key from the customers table and define it as the foreign key in the invoices table, thereby establishing the proper relationship between the tables.
When defining the relationships between tables, you must make the  **CONSTRAINT** declarations at the field level. This means that the constraints are defined within a **[CREATE TABLE](http://msdn.microsoft.com/library/FC45D36E-6E43-C030-5016-CCA8BB1379FE%28Office.15%29.aspx)** statement. To apply the constraints, use the **CONSTRAINT** keyword after a field declaration, name the constraint, name the table that it references, and name the field or fields within that table that will make up the matching foreign key.
The following statement assumes that the tblCustomers table has already been built, and that it has a primary key defined on the CustomerID field. The statement now builds the tblInvoices table, defining its primary key on the InvoiceID field. It also builds the one-to-many relationship between the tblCustomers and tblInvoices tables by defining another CustomerID field in the tblInvoices table. This field is defined as a foreign key that references the CustomerID field in the customers table. Note that the name of each constraint follows the  **CONSTRAINT** keyword.



```sql
CREATE TABLE tblInvoices  
    (InvoiceID INTEGER CONSTRAINT PK_InvoiceID PRIMARY KEY, 
    CustomerID INTEGER NOT NULL CONSTRAINT FK_CustomerID  
        REFERENCES tblCustomers (CustomerID),  
    InvoiceDate DATETIME, 
    Amount CURRENCY) 

```

Note that the primary key index (PK_InvoiceID) for the invoices table is declared within the  **CREATE TABLE** statement. To enhance the performance of the primary key, an index is automatically created for it, so there is no need to use a separate **[CREATE INDEX](http://msdn.microsoft.com/library/C5919EF4-A08D-DF06-7078-5331ADBCB45C%28Office.15%29.aspx)** statement.
Now create a shipping table that will contain each customer's shipping address. Assume that there will be only one shipping record for each customer record, so you will be establishing a one-to-one relationship.



```sql
CREATE TABLE tblShipping  
    (CustomerID INTEGER CONSTRAINT PK_CustomerID PRIMARY KEY 
        REFERENCES tblCustomers (CustomerID),  
    Address TEXT(50), 
    City TEXT(50), 
    State TEXT(2), 
    Zip TEXT(10)) 

```

Note that the CustomerID field is both the primary key for the shipping table and the foreign key reference to the customers table.

## Constraints

Constraints can be used to establish primary keys and referential integrity, and to restrict values that can be inserted into a field. In general, constraints can be used to preserve the integrity and consistency of the data in your database.

There are two types of constraints: a single-field or field-level constraint, and a multi-field or table-level constraint. Both kinds of constraints can be used in either the  **CREATE TABLE** or the **[ALTER TABLE](http://msdn.microsoft.com/library/78E6C92C-E88C-E55F-6B89-435360C166A6%28Office.15%29.aspx)** statement.

A single-field constraint, also known as a column-level constraint, is declared with the field itself, after the field and data type have been declared. Use the customers table and create a single-field primary key on the CustomerID field. To add the constraint, use the  **CONSTRAINT** keyword with the name of the field.




```sql
ALTER TABLE tblCustomers 
   ALTER COLUMN CustomerID INTEGER 
   CONSTRAINT PK_tblCustomers PRIMARY KEY 

```

Notice that the name of the constraint is given. You could use a shortcut for declaring the primary key that omits the  **CONSTRAINT** clause entirely.




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

Constraints also can be used to restrict the allowable values for a field. You can restrict values to  **NOT NULL** or **UNIQUE**, or you can define a check constraint, which is a type of business rule that can be applied to a field. Assume that you want to restrict (or constrain) the values of the first name and last name fields to be unique, meaning that there should never be a combination of first name and last name that is the same for any two records in the table. Because this is a multi-field constraint, it is declared at the table level, not the field level. Use the **ADD CONSTRAINT** clause and define a multi-field list.




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

The expression used to define a check constraint also can refer to more than one field in the same table, or to fields in other tables, and can use any operations that are valid in Access SQL, such as  **[SELECT](http://msdn.microsoft.com/library/A5C9DA94-5F9E-0FC0-767A-4117F38A5EF3%28Office.15%29.aspx)** statements, mathematical operators, and aggregate functions. The expression that defines the check constraint can be no more than 64 characters long.

Suppose that you want to check each customer's credit limit before he or she is added to the customers table. Using an  **ALTER TABLE** statement with the **ADD COLUMN** and **CONSTRAINT** clauses, create a constraint that will look up the value in the CreditLimit table to verify the customer's credit limit. Use the following SQL statements to create the tblCreditLimit table, add the CustomerLimit field to the tblCustomers table, add the check constraint to the tblCustomers table, and test the check constraint.




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

Note that when you execute the  **[UPDATE TABLE](http://msdn.microsoft.com/library/08F9C3D6-C020-ECF1-5748-43B93A76DFBB%28Office.15%29.aspx)** statement, you receive a message indicating that the update did not succeed because it violated the check constraint. If you update the CustomerLimit field to a value that is equal to or less than 100, the update will succeed.


