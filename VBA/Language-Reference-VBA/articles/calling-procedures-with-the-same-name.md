---
title: Calling Procedures with the Same Name
keywords: vbcn6.chm1076672
f1_keywords:
- vbcn6.chm1076672
ms.prod: office
ms.assetid: 5d310675-136b-58bb-29e2-ca09726b8ce0
ms.date: 06/08/2017
---


# Calling Procedures with the Same Name

You can call a [procedure](vbe-glossary.md) located in any [module](vbe-glossary.md) in the same [project](vbe-glossary.md) as the active module just as you would call a procedure in the active module. However, if two or more modules contain a procedure with the same name, you must specify a module name in the calling statement, as shown in the following example:


```vb
Sub Main() 
    Module1.MyProcedure 
End Sub
```


If you give the same name to two different procedures in two different projects, you must specify a project name when you call that procedure. For example, the following procedure calls the  `Main` procedure in the `MyModule` module in the `MyProject.vbp` project.




```vb
Sub Main() 
    [MyProject.vbp].[MyModule].Main 
End Sub
```


 **Note**  Different applications have different names for a project. For example, in Microsoft Access, a project is called a database (.mdb); in Microsoft Excel, it's a workbook (.xls).


## Tips for Calling Procedures




- If you rename a module or project, be sure to change the module or project name wherever it appears in calling [statements](vbe-glossary.md); otherwise, Visual Basic will not be able to find the called procedure. You can use the  **Replace** command on the **Edit** menu to find and replace text in a module.
    
- To avoid naming conflicts among referenced projects, give your procedures unique names so you can call a procedure without specifying a project or module.
    



