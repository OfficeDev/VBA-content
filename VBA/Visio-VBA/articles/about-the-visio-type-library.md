---
title: About the Visio Type Library
keywords: vis_sdr.chm81901860
f1_keywords:
- vis_sdr.chm81901860
ms.prod: visio
ms.assetid: 583b7622-7736-a661-5600-862ecbd9f522
ms.date: 06/08/2017
---


# About the Visio Type Library

Visio products include a type library that defines the objects, properties, methods, events, and constants that Visio exposes to Automation clients. To use the Visio type library, a development environment must reference it. The Visual Basic for Applications (VBA) project of a Visio document automatically references the Visio type library. In other development environments you must take appropriate steps to reference the library.

The names of the libraries your VBA project references are displayed in the  **Project/Library** list in the **Object Browser** in the Visual Basic Editor.

## Benefits of using a type library

A type library is useful for the following reasons.


- The information in a type library serves as input to object browsers supplied by VBA and other development environments. You can use object browsers to view descriptions of objects supplied by Automation servers (such as the Visio application) installed on your system. For example, you can view the syntax of a Visio property, method, or event and paste code shown by the browser into your program.
    
- A type library allows development environments to bind your program's code to Automation server code at compile (design) time rather than dynamically at run time. The result is that your program often runs faster. For example, you can use  **Visio.Page**,  **Visio.Shape**,  **Visio.Document**, and so on instead of  **Object**.
    

## Resolving object name ambiguities

Your VBA project or Visual Basic program can reference many type libraries. Libraries sometimes declare items with the same name. For example, both Visio and Excel expose an object called  **Application**. When more than one library declares an item with the same name, VBA and Visual Basic bind the name to the library with the highest priority.

One way to resolve name ambiguities is to prefix object types with the corresponding library name. For example: 




```vb
Dim vsoApplication As Visio.Application 
Dim xlApplication As Excel.Application

```

If your code runs exclusively in the context of a VBA project of a Visio document, you don't have to prefix names of Visio object types with  _Visio_, although it is a good idea. If you do this, the Visio type library has a higher priority than other libraries that may declare conflicting names. VBA does not let you change the priority of the Visio type library when you are using VBA within Visio, but in other development environments you can change the priority of the Visio type library.


