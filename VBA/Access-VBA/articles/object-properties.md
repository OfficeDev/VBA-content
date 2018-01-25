---
title: Object Properties
keywords: vbaac10.chm5187733
f1_keywords:
- vbaac10.chm5187733
ms.prod: access
ms.assetid: 9fc87446-68bd-d592-71c8-8d8c022af2c4
ms.date: 06/08/2017
---


# Object Properties

  

**Applies to:** Access 2013 | Access 2016

 **In this article**
[Setting](#sectionSection0)
[Remarks](#sectionSection1)
[Example](#sectionSection2)


The  **Object** properties provide general information about objects contained in the Navigation Pane.

 **Note**  The  **Object** properties are available for all objects in a Microsoft Access database and only for forms, macros, modules, and reports in an Access project (.adp).


## Setting
<a name="sectionSection0"> </a>

You can view the  **Object** properties, and set the **Description** or **Attributes** properties, in the following ways:


- Click an object in the Database window. On the  **Database Tools** tab, in the **Show/Hide** group, click **Property Sheet**.
    
- Right-click an object in the Database window, and then click  **Properties** on the shortcut menu.
    
You can also specify or determine the  **Object** properties in an Access database by using Visual Basic . The **Object** properties of an Access project (.adp) are not available using Visual Basic.


 **Note**  You can only enter or edit the  **Description** and **Attributes** properties. The other **Object** properties are set by Microsoft Access and are read-only.


## Remarks
<a name="sectionSection1"> </a>

The objects for which you can display properties in the Database window are tables, queries, forms, reports, macros, and modules. Each class of objects in the database is represented by a separate DAO  **Document** object within the DAO **Containers** collection. For example, the **Containers** collection contains a **Document** object that represents all the forms in the database.

The following  **Object** properties are available from the Database window.



|**Property**|**Description**|
|:-----|:-----|
|**Name**|This is the name of the object and contains the setting from the object's  **Name** property.|
|**Type**|This is the object's type. Microsoft Access object types are Form, Macro, Module, Query, Report, and Table.|
|**Description**|This is the object's description and is the same as the setting for the object's  **Description** property. You can also set the object's **Description** property in the object's property sheet.|
|**Created**|This is the date that the object was created. For tables and queries, this property is the same as the  **DateCreated** property.|
|**Modified**|This is the date that the object was last modified. For tables and queries, this property is the same as the  **LastUpdated** property.|
|**Owner**|This is the owner of the object. For more information, see the  **Owner** property.|
|**Attributes**|This property specifies whether the object is hidden or visible and whether the object can be replicated in a database replica. If you set the Hidden attribute to  **True** (by selecting the **Hidden** check box), the object won't appear in the Database window. To display hidden objects in the Navigation Pane, click the **Microsoft Office Button** ![File menu button](images/O12FileMenuButton_ZA10077102.gif) and then click **Access Options**. Click the  **Current Database** category, and then click **Navigation Options**. Click  **Show Hidden Objects** and then click **OK**. The icons for hidden objects will be dimmed in the Database window. You can then turn the Hidden attribute off, making the objects visible in the Database window.|

## Example
<a name="sectionSection2"> </a>

The following example uses the PrintObjectProperties subroutine to print the values of an object's  **Object** properties to the Debug window. The subroutine requires the object type and object name as arguments.


```vb
Dim strObjectType As String 
Dim strObjectName As String 
Dim strMsg As String 
 
strMsg = "Enter object type (e.g., Forms, Scripts, " _ 
 &; "Modules, Reports, Tables)." 
' Get object type. 
strObjectType = InputBox(strMsg) 
strMsg = "Enter the name of a form, macro, module, " _ 
 &; "query, report, or table." 
' Get object name from user. 
strObjectName = InputBox(strMsg) 
' Pass object type and object name to 
' PrintObjectProperties subroutine. 
PrintObjectProperties strObjectType, strObjectName 
 
Sub PrintObjectProperties(strObjectType As String, strObjectName _ 
 As String) 
Dim dbs As Database, ctr As Container, doc As Document 
Dim intI As Integer 
Dim strTabChar As String 
Dim prp As DAO.Property 
 
Set dbs = CurrentDb 
strTabChar = vbTab 
' Set Container object variable. 
Set ctr = dbs.Containers(strObjectType) 
' Set Document object variable. 
Set doc = ctr.Documents(strObjectName) 
doc.Properties.Refresh 
' Print the object name to Debug window. 
Debug.Print doc.Name 
' Print each Object property to Debug window. 
For Each prp in doc.Properties 
 Debug.Print strTabChar &; prp.Name &; " = " &; prp.Value 
Next 
End Sub
```

 **ACCESS SUPPORT RESOURCES**<br>
[Access for developers forum on MSDN](https://social.msdn.microsoft.com/Forums/office/en-US/home?forum=accessdev)<br>
[Access help on support.office.com](https://support.office.com/search/results?query=Access)<br>
[Access help on answers.microsoft.com](http://answers.microsoft.com/en-us/office/forum/access?page=1&;tab=question&;status=all&;auth=1)<br>
[Search for specific Access error codes on Bing](http://www.bing.com/)<br>
[Access forums on UtterAccess](http://www.utteraccess.com/forum/index.php?act=idx)<br>
[Access wiki on UtterAcess](http://www.utteraccess.com/forum/index.php?act=idx)<br>
[Access developer and VBA programming help center (FMS)](http://www.fmsinc.com/MicrosoftAccess/developer/)<br>
[Access posts on StackOverflow](http://stackoverflow.com/questions/tagged/ms-access)

