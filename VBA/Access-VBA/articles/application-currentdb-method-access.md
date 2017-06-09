---
title: Application.CurrentDb Method (Access)
keywords: vbaac10.chm12546
f1_keywords:
- vbaac10.chm12546
ms.prod: access
api_name:
- Access.Application.CurrentDb
ms.assetid: defcf58f-7689-90e0-001c-ba5e7e87eb88
ms.date: 06/08/2017
---


# Application.CurrentDb Method (Access)

The  **CurrentDb** method returns an object variable of type **Database** that represents the database currently open in the Microsoft Access window.


## Syntax

 _expression_. **CurrentDb**

 _expression_ A variable that represents an **Application** object.


### Return Value

Database


## Remarks




 **Note**  In Microsoft Access the  **CurrentDb** method establishes a hidden reference to the Microsoft Office 12.0 Access Conectivity Engine object library in a Microsoft Access database.

In order to manipulate the structure of your database and its data from Visual Basic, you must use Data Access Objects (DAO). The  **CurrentDb** method provides a way to access the current database from Visual Basic code without having to know the name of the database. Once you have a variable that points to the current database, you can also access and manipulate other objects and collections in the DAO hierarchy.

You can use the  **CurrentDb** method to create multiple object variables that refer to the current database. In the following example, the variables `dbsA` and `dbsB` both refer to the current database:




```vb
Dim dbsA As Database, dbsB As Database 
Set dbsA = CurrentDb 
Set dbsB = CurrentDb
```


 **Note**  In previous versions of Microsoft Access, you may have used the syntax  `DBEngine.Workspaces(0).Databases(0)`or  `DBEngine(0)(0)`to return a pointer to the current database. In Microsoft Access 2000, you should use the  **CurrentDb** method instead. The **CurrentDb** method creates another instance of the current database, while the `DBEngine(0)(0)`syntax refers to the open copy of the current database. The  **CurrentDb** method enables you to create more than one variable of type **Database** that refers to the current database. Microsoft Access still supports the `DBEngine(0)(0)`syntax, but you should consider making this modification to your code in order to avoid possible conflicts in a multiuser database.

If you need to work with another database at the same time that the current database is open in the Microsoft Access window, use the  **OpenDatabase** method of a **Workspace** object. The **OpenDatabase** method doesn't actually open the second database in the Microsoft Access window; it simply returns a **Database** variable representing the second database. The following example returns a pointer to the current database and to a database called Contacts.mdb:




```vb
Dim dbsCurrent As Database, dbsContacts As Database 
Set dbsCurrent = CurrentDb 
Set dbsContacts = DBEngine.Workspaces(0).OpenDatabase("Contacts.mdb")
```


## See also


#### Concepts


[Application Object](application-object-access.md)

