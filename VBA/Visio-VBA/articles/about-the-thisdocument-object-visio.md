---
title: About the ThisDocument Object (Visio)
keywords: vis_sdr.chm1059157
f1_keywords:
- vis_sdr.chm1059157
ms.prod: visio
ms.assetid: da3df7b4-3eaf-2603-1a1e-2ed737eb1d43
ms.date: 06/08/2017
---


# About the ThisDocument Object (Visio)

TheVisual Basic for Applications (VBA) project of every Visio document has a class module called  **ThisDocument**. When referenced from code in the project, the  **ThisDocument** object returns a reference to the project's **Document** object.

You can display the name of the VBA project's document in a message box by using the following statement, for example: 



```vb
MsgBox ThisDocument.Name
```

You can get the first page of the VBA project's document by using the following code, for example: 



```vb
Dim vsoPage As Visio.Page 
Set vsoPage = ThisDocument.Pages.Item(1)
```

If you want to manipulate the document associated with your VBA project, use the  **ThisDocument** object. If you want to manipulate a document, but not necessarily the document associated with your VBA project, get a **Document** object from the **Documents** collection.
The  **ActiveDocument** property often, but not necessarily, returns a reference to the same document as the **ThisDocument** object. The **ActiveDocument** and **ThisDocument** objects are the same if the document shown in the Visio active window is the document containing the **ThisDocument** object's project. Whether your code uses the **ActiveDocument** or **ThisDocument** object depends on the purpose of your program.

 **Note**  You can extend the set of properties and methods of a project's  **Document** object by adding public properties and methods to that project's **ThisDocument** class module. The new methods and properties are exposed just like the built-in methods and properties implemented by Visio. The new methods and properties are not available when you reference other **Document** objects. The **ThisDocument** object is not available to code that is not part of the VBA project of a Visio document.


