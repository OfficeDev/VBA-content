---
title: VisWebPageSettings Object (Visio Save As Web)
ms.prod: visio
ms.assetid: 1f286540-2c46-4a2a-b133-2bfd6168db36
ms.date: 06/08/2017
---


# VisWebPageSettings Object (Visio Save As Web)

Contains the settings for the web page.


## Remarks

The  **VisWebPageSettings** object serves as a container for a web page's properties.

Many of the properties of the  **VisWebPageSettings** object correspond to the settings available in the **Save As** dialog box when a user clicks the **File** tab, clicks **Export**, clicks  **Change File Type**, clicks  **Web Page (*.htm)**, and then clicks  **Save As**.

For example, the  **[PageTitle](viswebpagesettings-pagetitle-property-visio-save-as-web.md)** property, which contains the title that appears in the title bar when a web page is displayed in a browser, corresponds to the value in the **Page title** box in the **Set Page Title** dialog box (in the **Save As** dialog box, click **Change Title**). Also, the  **[DispScreenRes](viswebpagesettings-dispscreenres-property-visio-save-as-web.md)** property corresponds to the value selected in the **Target Monitor** list on the **Advanced** tab of the **Save As Web Page** dialog box (in the **Save As** dialog box, in the **Save as type** list, select **Web Page (*.htm;*.html)**, and then click  **Publish**).

When you want to create a web page, use the  **[WebPageSettings](vissaveasweb-webpagesettings-property-visio-save-as-web.md)** property of the ** [VisSaveAsWeb](http://msdn.microsoft.com/library/c4675de8-0f63-179f-f687-8962d54d6b2f%28Office.15%29.aspx)** object to get a reference to the **VisWebPageSettings** object, which you can use to set the web page's properties, as shown in the following example.




```vb
Public Sub VisWebPageSettingsObject_Example() 
 Dim vsoWebSettings As VisWebPageSettings 
 Dim vsoSaveAsWeb As VisSaveAsWeb 
 
 ' Query Visio for the VisSaveAsWeb object. 
 Set vsoSaveAsWeb = Visio.Application.SaveAsWebObject 
 
 ' Get a WebPageSettings object. 
 Set vsoWebSettings = vsoSaveAsWeb.WebPageSettings 
 
 With vsoWebSettings 
 
 ' Set the title that is displayed in the browser's title bar. 
 .PageTitle = "AccountingDeptOrgChart082501" 
 
 ' Prevent dialog boxes from appearing in the user interface. 
 .QuietMode = True 
 End With 
 
 ' If you do not call the AttachToVisioDoc method to 
 ' identify a specific document, Visio saves the 
 ' active document by default. 
 vsoSaveAsWeb.CreatePages 
End Sub
```


 **Note**  To view the  **VisWebPageSettings** class in the Object Browser, make sure that you have a reference to the Save As Web Page DLL in your project (in the Visual Basic Editor window, click **References** on the **Tools** menu, and then select the **Microsoft Visio 15`.0 Save As Web Type Library** check box in the **Available References** list).


