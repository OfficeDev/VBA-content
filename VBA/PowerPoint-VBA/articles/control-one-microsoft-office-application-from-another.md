---
title: Control One Microsoft Office Application from Another
keywords: vbapp10.chm5192317
f1_keywords:
- vbapp10.chm5192317
ms.prod: powerpoint
ms.assetid: 435990be-0ce5-8a7b-4e5e-c4a5e7396524
ms.date: 06/08/2017
---


# Control One Microsoft Office Application from Another

If you want to run code in one Office application that works with the objects in another application, follow these steps.


1. Set a reference to the other application's type library in the  **References** dialog box ( **Tools** menu). After you have done this, the objects, properties, and methods show up in the Object Browser and the syntax is checked at compile time. You can also get context-sensitive Help on them.
    
2. Declare object variables that will refer to the objects in the other application as specific types. Make sure you qualify each type with the name of the application that is supplying the object. For example, the following statement declares a variable that will point to a Word document, and another that refers to an Excel application. 
    
```vb
Dim appWD As Word.Application, wbXL As Excel.Application 
  ```


     ** Note** You must follow the steps above if you want your code to be early bound.
    
3. Use the  **New** keyword with the OLE Programmatic Identifier of the object you want to work with in the other application, as shown in the following example. If you want to see the session of the other application, set the **Visible** property to **True**.
    
```vb
Dim appWD As Word.Application  
Set appWD = New Word.Application  
appWd.Visible = True 
  ```

4. Apply properties and methods to the object contained in the variable. For example, the following instruction creates a Word document. 
    
```vb
Dim appWD As Word.Application 
Set appWD = New Word.Application 
appWD.Documents.Add 
  ```

5. When you have finished working with the other application, use the  **Quit** method to close it, as shown in the following example.
    
  ```
  appWd.Quit 
  ```


