---
title: About Extending the Functionality of Visio
keywords: vis_sdr.chm81901861
f1_keywords:
- vis_sdr.chm81901861
ms.prod: visio
ms.assetid: ddd8ce40-7df3-67ef-7365-9f728b3a8c39
ms.date: 06/08/2017
---


# About Extending the Functionality of Visio

You can extend the functionality of Microsoft Visio in the following ways:


- Create Visio-specific macros and add-ons.
    
- Create COM (Component Object Model) add-ins.
    

## Macros and add-ons

Macros and add-ons are programs that extend the functionality of Visio. Exactly how you run a macro or add-on depends on the context for which it was designed.

You can run a macro or add-on from the Visio application in several ways. Here are a few of the most common:


- Choose a macro or add-on from the  **Macros** dialog box. (In the **Code** group on the [Developer](run-visio-in-developer-mode.md) tab, click **Macros**.) If your program is an EXE file, before it can appear in the  **Macros** dialog box, it must be located in a folder along the **Add-ons** path in the **File Locations** dialog box. (Click the **File** tab, click **Options**, click  **Advanced**, and then, under  **General**, click ** File Locations**.) 
    
     **Note**  Starting with Microsoft Office Visio 2003, instead of specifying file paths to your Visio add-ons, you can publish your add-ons by using a Microsoft Windows Installer package. By doing so, you can take advantage of Microsoft Office application features such as language switching, installation on demand, and repair. For more information about using a Windows Installer package to publish your add-ons, search for "Microsoft Windows Installer" on MSDN, the Microsoft Developer Network Web site.
- Double-click a shape associated with an add-on or macro. The program you want to run must be selected in the  **Run macro** list on the **Double-Click** tab in the **Behavior** dialog box for that shape. (Select the shape, and then, on the [Developer](run-visio-in-developer-mode.md) tab, click **Behavior**).
    
- Right-click a shape, and then click a custom menu item for an add-on or macro on the shortcut menu. The program associated with the custom menu item must be entered in the Actions section of the ShapeSheet window for the shape.
    
If an add-on is designed to be run outside the Visio application, you run it like any Microsoft Windows-based program (for example, by double-clicking an icon on the desktop). For details, see your Windows documentation.


## COM add-ins

Beginning with Visio 2002, you can use COM add-ins in the same standardized way as in other Microsoft Office applications. The COM add-in must be registered with the Visio application and can work in multiple applications. For example, you can build a COM add-in that performs the same task in Visio and Microsoft Excel, or any of the Microsoft Office applications. You can create COM add-ins with Microsoft Visual Basic 5.0 and higher, Microsoft C++, Microsoft Office 2000 Developer Edition and higher, or any of the Microsoft Visual Studio .NET applications.

For more information about building COM add-ins, see MSDN.


