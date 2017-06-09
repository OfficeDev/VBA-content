---
title: InvisibleApp.SaveAsWebObject Property (Visio)
keywords: vis_sdr.chm17551660
f1_keywords:
- vis_sdr.chm17551660
ms.prod: visio
api_name:
- Visio.InvisibleApp.SaveAsWebObject
ms.assetid: 9020e7db-b696-7484-c024-fd92906e486b
ms.date: 06/08/2017
---


# InvisibleApp.SaveAsWebObject Property (Visio)

Returns a reference to the  **IDispatch** interface of a **VisSaveAsWeb** object. Read-only.


## Syntax

 _expression_ . **SaveAsWebObject**

 _expression_ A variable that represents an **InvisibleApp** object.


### Return Value

Object


## Remarks

Once you have a reference to the  **VisSaveAsWeb** object, you can use the objects, methods, and properties of the Save as Web Page API to publish Microsoft Visio documents to the Web. For more information about the Save as Web Page API, search for "Save as Web Page API" on MSDN.

To be able to work with the Save as Web Page API, you must get a reference to the  **Microsoft Visio 14.0 Save As Web Type Library** in your Microsoft Visual Basic or Microsoft Visual Basic for Applications (VBA) project. To get this reference in VBA, use the following procedure:


1. In the  **Code** group on the[Developer](http://msdn.microsoft.com/library/1bdc55f5-8fc7-7257-03d5-c049eceb29ff%28Office.15%29.aspx) tab, click **Visual Basic**.
    
2. On the  **Tools** menu, click **References**.
    
3. In the  **Available References** list, select **Microsoft Visio 14.0 Save As Web Type Library** and click **OK**.
    

## Example

This VBA macro shows how to use the  **SaveAsWebObject** property to get a **VisSaveAsWeb** object. It also shows how to get a **VisWebPageSettings** object, configure Web-page settings, and create a Web page to display the active Visio document. The macro gets a Visio **Application** object and passes it to the **SaveAsWeb** procedure, which gets the **VisSaveAsWeb** object, configures the settings, and creates the Web page.

Before running this macro, get a reference to the  **Microsoft Visio 14.0 Save As Web Type Library** as described above, and replace _path\filename_ in the code with the full path to and name of the .htm file you want to create on your computer to display the Web page.




```vb
 
Public Sub SaveAsWebObject_Example 
 
 Dim vsoApplication as Visio.Application 
 Call SaveAsWeb(vsoApplication) 
 
End Sub 
 
 
Public Sub SaveAsWeb (vsoApplication as Visio.Application) 
 
 Dim objSaveAsWeb As IVisSaveAsWeb 
 Dim objWebPageSettings As IVisWebPageSettings 
 
 ' Get a VisSaveAsWeb object that 
 ' represents a new Web page project 
 Set objSaveAsWeb = Application.SaveAsWebObject 
 
 ' Get a VisWebPageSettings object 
 Set objWebPageSettings = objSaveAsWeb.WebPageSettings 
 
 ' Configure Web-page settings 
 objWebPageSettings.StartPage = 1 
 objWebPageSettings.EndPage = 2 
 objWebPageSettings.LongFileNames = True 
 objWebPageSettings.TargetPath = "path\filename " 
 
 ' Now create the pages; because we did not identify 
 ' a particular document, the active document is saved 
 objSaveAsWeb.CreatePages 
 
End Sub
```


