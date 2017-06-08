---
title: CurrentProject Object (Access)
keywords: vbaac10.chm12739
f1_keywords:
- vbaac10.chm12739
ms.prod: access
api_name:
- Access.CurrentProject
ms.assetid: e6baae73-1eeb-b48f-d35e-b3e921378561
ms.date: 06/08/2017
---


# CurrentProject Object (Access)

The  **CurrentProject** object refers to the project for the current Microsoft Access project (.adp) or Access database.


## Remarks

The  **CurrentProject** object has several collections that contain specific **[AccessObject](accessobject-object-access.md)** objects within the current database. The following table lists the name of each collection and the types of objects it contains.



|**Collections**|**Object type**|
|:-----|:-----|
|**[AllForms](allforms-object-access.md)**|All forms|
|**[AllReports](http://msdn.microsoft.com/library/5846cf60-41b4-e9f8-ea27-b9400a6d3861%28Office.15%29.aspx)**|All reports|
|**[AllMacros](http://msdn.microsoft.com/library/a36ba978-f643-aca6-5efb-842723d17bbc%28Office.15%29.aspx)**|All macros|
|**[AllModules](http://msdn.microsoft.com/library/322815ae-3afd-f299-0ce9-2e9dbbb8536a%28Office.15%29.aspx)**|All modules|

 **Note**  The collections in the preceding table contain all of the respective objects in the database regardless if they are opened or closed.

For example, an  **AccessObject** object representing a form is a member of the **AllForms** collection, which is a collection of **AccessObject** objects within the current database. Within the **AllForms** collection, individual members of the collection are indexed beginning with zero. You can refer to an individual **AccessObject** object in the **AllForms** collection either by referring to the form by name, or by referring to its index within the collection. If you want to refer to a specific object in the **AllForms** collection, it's better to refer to it by name because a item's collection index may change. If the object name includes a space, the name must be surrounded by brackets ([ ]).



|**Syntax**|**Example**|
|:-----|:-----|
|**AllForms** ! _formname_|AllForms!OrderForm|
|**AllForms** ![ _form name_]|AllForms![Order Form]|
|**AllForms** (" _formname_")|AllForms("OrderForm")|
|**AllForms** ( _formname_)|AllForms(0)|

## Example

The following example prints some current property settings of the  **CurrentProject** object and then sets an option to display hidden objects within the application:


```
Sub ApplicationInformation() 
 ' Print name and type of current object. 
 Debug.Print Application.CurrentProject.FullName 
 Debug.Print Application.CurrentProject.ProjectType 
 ' Set Hidden Objects option under Show on View Tab 
 'of the Options dialog box. 
 Application.SetOption "Show Hidden Objects", True 
End Sub
```

The next example shows how to use the CurrentProject object using Automation from another Microsoft Office application. First, from the other application, create a reference to Microsoft Access by clicking  **References** on the **Tools** menu in the Module window. Select the check box next to **Microsoft Access Object Library**. Then enter the following code in a Visual Basic module within that application and call the GetAccessData procedure.

The example passes a database name and report name to a procedure that creates a new instance of the  **Application** class, opens the database, and verifies that the specified report exists using the **CurrentProject** object and **AllReports** collection.




```
Sub GetAccessData() 
' Declare object variable in declarations section of a module 
 Dim appAccess As Access.Application 
 Dim strDB As String 
 Dim strReportName As String 
 
 strDB = "C:\Program Files\Microsoft "_ 
 &amp; "Office\Office11\Samples\Northwind.mdb" 
 strReportName = InputBox("Enter name of report to be verified", _ 
 "Report Verification") 
 VerifyAccessReport strDB, strReportName 
End Sub 
 
Sub VerifyAccessReport(strDB As String, _ 
 strReportName As String) 
 ' Return reference to Microsoft Access 
 ' Application object. 
 Set appAccess = New Access.Application 
 ' Open database in Microsoft Access. 
 appAccess.OpenCurrentDatabase strDB 
 ' Verify report exists. 
 On Error Goto ErrorHandler 
 appAccess.CurrentProject.AllReports(strReportName) 
 MsgBox "Report " &amp; strReportName &amp; _ 
 " verified within Northwind database." 
 appAccess.CloseCurrentDatabase 
 Set appAccess = Nothing 
Exit Sub 
ErrorHandler: 
 MsgBox "Report " &amp; strReportName &amp; _ 
 " does not exist within Northwind database." 
 appAccess.CloseCurrentDatabase 
 Set appAccess = Nothing 
End Sub
```


## Methods



|**Name**|
|:-----|
|[AddSharedImage](http://msdn.microsoft.com/library/c6c02f12-6c5f-852a-65b7-a0ffbb3346fd%28Office.15%29.aspx)|
|[CloseConnection](http://msdn.microsoft.com/library/f2feac44-e509-48d7-e815-e0cf2935d7b9%28Office.15%29.aspx)|
|[OpenConnection](http://msdn.microsoft.com/library/37b5d50c-ddc9-97d4-2b8f-068ba2702e6d%28Office.15%29.aspx)|
|[UpdateDependencyInfo](http://msdn.microsoft.com/library/90461646-22a6-bfa8-4663-9f05c8ac3757%28Office.15%29.aspx)|

## Properties



|**Name**|
|:-----|
|[AccessConnection](http://msdn.microsoft.com/library/c2bf2846-c5ab-34a2-4b24-33c9cc9820c4%28Office.15%29.aspx)|
|[AllForms](http://msdn.microsoft.com/library/4933a409-0d15-16ee-69a3-d78b0f2685c7%28Office.15%29.aspx)|
|[AllMacros](http://msdn.microsoft.com/library/73c01f69-530b-eb7f-8f77-ecf47e9c2d2f%28Office.15%29.aspx)|
|[AllModules](http://msdn.microsoft.com/library/2d6f5786-c431-9c1a-b581-56fb969fb947%28Office.15%29.aspx)|
|[AllReports](http://msdn.microsoft.com/library/dda91007-88ef-5660-f67f-4cc9c6f5dbb3%28Office.15%29.aspx)|
|[Application](http://msdn.microsoft.com/library/565628df-7dbc-be17-9c8a-80de222a1583%28Office.15%29.aspx)|
|[BaseConnectionString](http://msdn.microsoft.com/library/280bb905-d321-d844-8ab6-6c9352dd3ab0%28Office.15%29.aspx)|
|[Connection](http://msdn.microsoft.com/library/ab956942-deff-793f-e5e6-7412554f9950%28Office.15%29.aspx)|
|[FileFormat](http://msdn.microsoft.com/library/eb062d95-3042-eae7-9c0b-9d052e28b8cd%28Office.15%29.aspx)|
|[FullName](http://msdn.microsoft.com/library/43fa4260-4e70-c314-c02d-1328b7c1b2a2%28Office.15%29.aspx)|
|[ImportExportSpecifications](http://msdn.microsoft.com/library/b614eb40-d9cd-d615-41c9-c6980ea85006%28Office.15%29.aspx)|
|[IsConnected](http://msdn.microsoft.com/library/04e1123b-ad18-9ebc-3dec-f49bcc16d5a0%28Office.15%29.aspx)|
|[IsTrusted](http://msdn.microsoft.com/library/c3d8b6f8-c79f-79ab-d4e0-0454f97ac937%28Office.15%29.aspx)|
|[IsWeb](http://msdn.microsoft.com/library/dbcd7b51-75d1-54c7-9c49-7b1ea403c4d9%28Office.15%29.aspx)|
|[Name](http://msdn.microsoft.com/library/b7eb012e-6145-d962-8884-3ccf3eaf46fd%28Office.15%29.aspx)|
|[Parent](http://msdn.microsoft.com/library/afd66c1b-db13-e336-02db-fcdc8f5226bc%28Office.15%29.aspx)|
|[Path](http://msdn.microsoft.com/library/25f28502-b5fc-aafa-9189-eb091907a529%28Office.15%29.aspx)|
|[ProjectType](http://msdn.microsoft.com/library/b68e5888-0bea-ae7a-b389-b87c7002352c%28Office.15%29.aspx)|
|[Properties](http://msdn.microsoft.com/library/fd53f73f-184a-0793-da0d-7bcd95c20439%28Office.15%29.aspx)|
|[RemovePersonalInformation](http://msdn.microsoft.com/library/d98f2b2a-304f-8d27-14ad-55407f335f1e%28Office.15%29.aspx)|
|[Resources](http://msdn.microsoft.com/library/2edc7258-77b3-5d09-22eb-1620d460f0f3%28Office.15%29.aspx)|
|[WebSite](http://msdn.microsoft.com/library/ab2cc5f8-bd24-9f88-2598-1d8e6c71895e%28Office.15%29.aspx)|
|[IsSQLBackend](http://msdn.microsoft.com/library/39e312e0-9b58-e1fe-7a98-be5e225a3c0c%28Office.15%29.aspx)|

## See also


#### Other resources


[Access Object Model Reference](http://msdn.microsoft.com/library/2de134a4-6c5c-d2a3-8377-f4dd973ba650%28Office.15%29.aspx)
[CurrentProject Object Members](http://msdn.microsoft.com/library/adb319f1-487a-d7d1-5755-d57c31c776b8%28Office.15%29.aspx)
