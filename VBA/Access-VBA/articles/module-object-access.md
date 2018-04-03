---
title: Module Object (Access)
keywords: vbaac10.chm12268
f1_keywords:
- vbaac10.chm12268
ms.prod: access
api_name:
- Access.Module
ms.assetid: e04272fa-9c29-2567-bd15-1cea38906894
ms.date: 06/08/2017
---


# Module Object (Access)

A  **Module** object refers to a standard module or a class module.


## Remarks

Microsoft Access includes class modules that are not associated with any object, and form modules and report modules, which are associated with a form or report.

To determine whether a  **Module** object represents a standard module or a class module from code, check the **Module** object's **Type** property.

The  **[Modules](http://msdn.microsoft.com/library/f60a9929-4b79-cfed-8fb3-a4869a3afe9f%28Office.15%29.aspx)** collection contains all open **Module** objects, regardless of their type. Modules in the **Modules** collection can be compiled or uncompiled.

To return a reference to a particular standard or class  **Module** object in the **Modules** collection, use any of the following syntax forms.



|**Syntax**|**Description**|
|:-----|:-----|
|**Modules** !modulename|The  _modulename_ argument is the name of the **Module** object.|
|**Modules** (" _modulename_")|The  _modulename_ argument is the name of the **Module** object.|
|**Modules** ( _index_)|The  _index_ argument is the numeric position of the object within the collection.|
The following example returns a reference to a standard  **Module** object and assigns it to an object variable:




```
Dim mdl As Module 
Set mdl = Modules![Utility Functions]
```

Note that the brackets enclosing the name of the  **Module** object are necessary only if the name of the **Module** includes spaces.

The next example returns a reference to a form  **Module** object and assigns it to an object variable:




```
Dim mdl As Module 
Set mdl = Modules!Form_Employees
```

To refer to a specific form or report module, you can also use the  **[Form](http://msdn.microsoft.com/library/72ef9219-142b-b690-b696-3eba9a5d4522%28Office.15%29.aspx)** or **[Report](report-object-access.md)** object's **Module** property:




```
Forms!formname .Module
```

The following example also returns a reference to the  **Module** object associated with an Employees form and assigns it to an object variable:




```
Dim mdl As Module 
Set mdl = Forms!Employees.Module
```

Once you've returned a reference to a  **Module** object, you can set or read its properties and apply its methods.


## Methods



|**Name**|
|:-----|
|[AddFromFile](http://msdn.microsoft.com/library/a782b4dc-a4be-5166-3ce3-deb87ed1195b%28Office.15%29.aspx)|
|[AddFromString](http://msdn.microsoft.com/library/119db9d9-fac2-b86f-be21-c94366bda7d6%28Office.15%29.aspx)|
|[CreateEventProc](http://msdn.microsoft.com/library/13d2a4db-ec80-4225-f3fd-87527dbf660e%28Office.15%29.aspx)|
|[DeleteLines](http://msdn.microsoft.com/library/57f65c6c-4d9c-3abd-065b-b75d1ada06cb%28Office.15%29.aspx)|
|[Find](http://msdn.microsoft.com/library/6b8fcd1a-a490-19a0-1692-fb01f213c639%28Office.15%29.aspx)|
|[InsertLines](http://msdn.microsoft.com/library/54ea5ce3-fb2a-e9c7-85ef-8861141f63ec%28Office.15%29.aspx)|
|[InsertText](http://msdn.microsoft.com/library/105c77fe-29a3-ef93-3d01-8420f7725325%28Office.15%29.aspx)|
|[ReplaceLine](http://msdn.microsoft.com/library/9e267b4a-5c15-a1bc-e2e0-a528871c9268%28Office.15%29.aspx)|

## Properties



|**Name**|
|:-----|
|[Application](http://msdn.microsoft.com/library/9237a6d4-8c68-9d58-f696-6525f42963d0%28Office.15%29.aspx)|
|[CountOfDeclarationLines](http://msdn.microsoft.com/library/fc0bdb0f-264c-0311-946c-c5cdc03a86f0%28Office.15%29.aspx)|
|[CountOfLines](http://msdn.microsoft.com/library/6c3bb4c8-15a9-6365-155d-d28dc0c5de78%28Office.15%29.aspx)|
|[Lines](http://msdn.microsoft.com/library/a230ffef-6640-178f-b3a5-edd1e171a8f6%28Office.15%29.aspx)|
|[Name](http://msdn.microsoft.com/library/2bc2427a-5e0f-930c-e232-abfde3b0b614%28Office.15%29.aspx)|
|[Parent](http://msdn.microsoft.com/library/1d81a57b-bcc2-a8cb-8526-6fd6409d3131%28Office.15%29.aspx)|
|[ProcBodyLine](http://msdn.microsoft.com/library/b81affb6-a3ca-3bda-59f0-9fb809b34d2d%28Office.15%29.aspx)|
|[ProcCountLines](http://msdn.microsoft.com/library/d85cacb5-127a-68a1-3bff-cc13a8a7e9ed%28Office.15%29.aspx)|
|[ProcOfLine](http://msdn.microsoft.com/library/64a21820-923d-a816-6b6e-2a679d0e09ac%28Office.15%29.aspx)|
|[ProcStartLine](http://msdn.microsoft.com/library/ef9a1ab4-f992-5077-b52b-d16cba10f697%28Office.15%29.aspx)|
|[Type](http://msdn.microsoft.com/library/df30b007-5ce9-9de2-1013-747c47917288%28Office.15%29.aspx)|

## See also


#### Other resources


[Access Object Model Reference](http://msdn.microsoft.com/library/2de134a4-6c5c-d2a3-8377-f4dd973ba650%28Office.15%29.aspx)
[Module Object Members](http://msdn.microsoft.com/library/c2e71012-645e-b818-1247-9775f221619e%28Office.15%29.aspx)
