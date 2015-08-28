
# Create a DAO Recordset From a Form

 **Last modified:** July 28, 2015

 _**Applies to:** Access 2013_

You can create a  ** [Recordset](http://msdn.microsoft.com/library/9774232C-E6DA-175B-FC7F-ED2AB7908FA0%28Office.15%29.aspx)** object based on an Access form. To do so, use the ** [RecordsetClone](D73EF798-477D-9C36-6E29-82B22352C60B.md)** property of the form. This creates a dynaset-type **Recordset** that refers to the same underlying query or data as the form. If a form is based on a query, referring to the **RecordsetClone** property is the equivalent of creating a dynaset with the same query. You can use the **RecordsetClone** property when you want to apply a method that cannot be used with forms, such as the ** [FindFirst](http://msdn.microsoft.com/library/5FCF78CD-7D2C-2E47-14E5-996F2E14FF51%28Office.15%29.aspx)** method. The **RecordsetClone** property provides access to all the methods and properties that you can use with a dynaset.

The following example shows how to assign a  **Recordset** object to the records in the Orders form.



```
Dim rstOrders As DAO.Recordset 
 
Set rstOrders = Forms!Orders.RecordsetClone 

```

This code always creates the type of  **Recordset** being cloned (the type of **Recordset** on which the form is based); no other types are available. Note that the **Recordset** object is declared with the object library qualification. Because Access can use both DAO and ADO, it is better to fully qualify the data access variables by including the object library reference name.
