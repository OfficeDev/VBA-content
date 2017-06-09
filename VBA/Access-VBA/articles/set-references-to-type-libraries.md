---
title: Set References to Type Libraries
keywords: vbaac10.chm13780
f1_keywords:
- vbaac10.chm13780
ms.prod: access
ms.assetid: 6314a89b-89e9-d8c1-5964-889a361afcd1
ms.date: 06/08/2017
---


# Set References to Type Libraries

When you set a reference to another application's type library, you can use the objects supplied by that application in your code. For example, if you set a reference from Access to the Excel library, you can then use Excel objects through Automation (formerly called OLE Automation). If you set a reference to a Visual Basic project in another Access database, you can call its public procedures. If you set a reference to an ActiveX control, you can use that control on Access forms.

You can set a reference from Access while the Visual Basic Editor is open, or you can set a reference in Visual Basic code.

## Setting a Reference from Access

To set a reference to an application's type library:


1. On the  **Tools** menu, click **References**. The  **References** command on the **Tools** menu is available only when a Module window is open and active in Design view.
    
2. Select the check boxes for those applications whose type libraries you want to reference.
    

## Setting a Reference from Visual Basic

To set a reference from Visual Basic, you create a new  **Reference** object representing the desired reference. The **References** collection contains all currently set references.

To create a new  **Reference** object, use either the **AddFromFile** or **AddFromGUID** method of the **References** collection. To remove a **Reference** object, use the **Remove** method.


## Advantages of Setting References

Your Automation code will run faster if you set a reference to another application's type library before you work with its objects. If you've set a reference, you can declare an object variable representing an object in the other application as its most specific type. For example, if you're writing code to work with Excel objects, you can declare an object variable of type  **Excel.Application** by using the following syntax only if you've created a reference to the Excel type library:


```vb
Dim appXL As New Excel.Application
```

If you haven't set a reference to the Excel type library, you must declare the variable as a generic variable of type  **Object**. The following code runs more slowly:




```vb
Dim appXL As Object
```

Additionally, if you set a reference to an application's type library, all of its objects, as well as their methods and properties, are listed in the Object Browser. This makes it easy to determine what properties and methods are available to each object.

Since Access is an COM component that supports Automation, you can also set a reference to its type library from another application and work with Access objects from that application.


