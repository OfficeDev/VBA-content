---
title: Understanding Automation
keywords: vbcn6.chm1076677
f1_keywords:
- vbcn6.chm1076677
ms.prod: office
ms.assetid: 5b45f6f3-1459-ff25-51e1-32c475f11153
ms.date: 06/08/2017
---


# Understanding Automation

Automation (formerly OLE Automation) is a feature of the Component Object Model (COM), an industry-standard technology that applications use to expose their [objects](vbe-glossary.md) to development tools, macro languages, and other applications that support Automation. For example, a spreadsheet application may expose a worksheet, chart, cell, or range of cells — each as a different type of object. A word processor might expose objects such as an application, a document, a paragraph, a sentence, a bookmark, or a selection.

When an application supports Automation, the objects the application exposes can be accessed by Visual Basic. Use Visual Basic to manipulate these objects by invoking [methods](vbe-glossary.md) on the object or by getting and setting the object's properties. For example, you can create an [Automation object](vbe-glossary.md) named and write the following code to access the object:



```
MyObj.Insert "Hello, world." ' Place text. 
MyObj.Bold = True ' Format text. 
If Mac = True ' Check your platform constant 
 MyObj.SaveAs "HD:\WORDPROC\DOCS\TESTOBJ.DOC" ' Save the object (Macintosh). 
Else 
 MyObj.SaveAs "C:\WORDPROC\DOCS\TESTOBJ.DOC" ' Save the object (Windows). 

```

Use the following functions to access an Automation object:


|**Function**|**Description**|
|:-----|:-----|
|**CreateObject**|Creates a new object of a specified type.|
|**GetObject**|Retrieves an object from a file.|



For details on the properties and methods supported by an application, see the application documentation. The objects, functions, properties, and methods supported by an application are usually defined in the application's [object library](vbe-glossary.md).

