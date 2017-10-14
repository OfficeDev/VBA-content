---
title: "Using the Save as Web Page Object Model from Visual Basic: An Example"
ms.prod: visio
ms.assetid: c5833ff8-45f3-ab67-3b16-09c60238965a
ms.date: 06/08/2017
---


# Using the Save as Web Page Object Model from Visual Basic: An Example

To use the Save as Web Page API in your Visual Basic project, set a reference in your project to  **Microsoft Visio 15.0 Save As Web Type Library**.


 **Note**   In the Visual Basic Editor included with Visio, you can find the list of available references by clicking **References** on the **Tools** menu. In Visual Basic 6.0, you can find this list by clicking **References** on the **Project** menu.


The Save as Web Page model contains two classes:  **VisSaveAsWeb** and **VisWebPageSettings**, which implement the  **IVisSaveAsWeb** and **IVisWebPageSettings** interfaces, respectively.



- A  **VisSaveAsWeb** object implements the methods that perform the Web page creation process.
    
- A  **VisWebPageSettings** object contains the properties of your Web page project.
    

When you create a Web page and its supporting files (also called a Web page project), you'll typically follow these steps.


1. Use the  **SaveAsWebObject** property of the Visio **Application** object to get an instance of a **VisSaveAsWeb** object.
    
2. Use the  **WebPageSettings** property of the **VisSaveAsWeb** object to get a reference to a **VisWebPageSettings** object, which you can use to get or set the Web page settings for your project.
    
3. Set the properties of the  **VisWebPageSettings** object.
    
     **Note**  You must always provide a target path for your files.
4. Call the  **AttachToVisioDoc** method to identify the document to save as a Web page. If you don't specify which document to save, the active drawing is saved.
    
5. Call the  **CreatePages** method to begin the Save as Web Page operation.
    

The following procedure shows how to open a new Web page project, set selected properties, and create the Web page files.



```vb
Public Sub SaveAsWeb () 
 Dim vsoSaveAsWeb As VisSaveAsWeb 
 Dim vsoWebSettings As VisWebPageSettings 
 
 ' Get a VisSaveAsWeb object that 
 ' represents a new Web page project. 
 Set vsoSaveAsWeb = Visio.Application.SaveAsWebObject 
 
 ' Get a VisWebPageSettings object. 
 Set vsoWebSettings = vsoSaveAsWeb.WebPageSettings 
 
 ' Configure preferences. 
 With vsoWebSettings 
 .StartPage = 1 
 .EndPage = 2 
 .QuietMode = True 
 .TargetPath = "c:\your_folder_name\your_filename.htm" 
 End With 
 
 ' Create the pages. Because no particular document 
 ' is specified, the active drawing is saved. 
 vsoSaveAsWeb.CreatePages 
End Sub
```


