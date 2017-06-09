---
title: Check or Add an Object Library Reference
keywords: vbhw6.chm1107739
f1_keywords:
- vbhw6.chm1107739
ms.prod: office
ms.assetid: a04227a8-80e0-2eb3-52bb-f992d8bb5e68
ms.date: 06/08/2017
---


# Check or Add an Object Library Reference

If you use the objects in other applications as part of your Visual Basic application, you may want to establish a reference to the [object libraries](vbe-glossary.md) of those applications. Before you can do that, you must first be sure that the application provides an object library.

 **To see if an application provides an object library**




1. From the  **Tools** menu, choose **References** to display the **References** dialog box.
    
2. The  **References** dialog box shows all object libraries registered with the operating system. Scroll through the list for the application whose object library you want to reference. If the application isn't listed, you can use the **Browse** button to search for object libraries (*.olb and *.tlb) or[executable files](vbe-glossary.md) (*.exe and *.dll on Windows). References whose check boxes are checked are used by your[project](vbe-glossary.md); those that aren't checked are not used, but can be added.
    

 **To add a object library reference to your project**


- Select the object library reference in the  **Available References** box in the **References** dialog box and click **OK**. Your Visual Basic project now has a reference to the application's object library. If you open the[Object Browser](vbe-glossary.md) (press F2) and select the application's library, it displays the objects provided by the selected object library, as well as each object's[methods](vbe-glossary.md) and[properties](vbe-glossary.md). In the  **Object Browser**, you can select a[class](vbe-glossary.md) in the **Classes** box and select a method or property in the **Members** box. Use copy and paste to add the syntax to your code.
    


