---
title: Running Save as Web Page from the Command Line
ms.prod: visio
ms.assetid: 01dbf425-069f-5e11-0ace-5578c02c0b4b
ms.date: 06/08/2017
---


# Running Save as Web Page from the Command Line

The Save as Web Page feature is automatically installed with each Visio product. It is installed as a Visio add-on that has the name "SaveAsWeb."

To run the Save as Web Page feature from the command-line interface, you call the SaveAsWeb add-on and use the desired command-line options.

You can choose either of the following techniques:


- Create a formula that uses the RUNADDONWARGS function in a shape's event section. 
    
    You can do this in the ShapeSheet window without writing any code. For example, you could create a shape and insert a call to RUNADDONWARGS in the shape's double-click event. A user would need only to double-click the shape to create a Web page from the drawing. You can also use Automation to set formulas by using the  **Formula** property of the **Cell** object of the Visio object model.
    
    For details about the RUNADDONWARGS function,  **Cell** object, or **Formula** property, see the Visio Developer Reference (click **Help**, click  **Search**, and then click  **Developer Reference**). 
    
    For more details about using the RUNADDONWARGS function with Save as Web Page, see "Using the RUNADDONWARGS function" later in this topic.
    
- Write a Visual Basic macro in Visio (or write code in a separate component) that launches the SaveAsWeb add-on by using the Visio Automation object model. You can use the  **Run** method of the **Addon** object and pass the command-line parameters to specify the properties of the Web page.
    
    Using this technique may mean that you would write less code than if you used the Save as Web Page object model to specify parameters, but using the Run method requires familiarity with command-line parameters.
    
    For details about the  **Addon** object or **Run** method, see the Visio Developer Reference (click **Help**, click  **Search**, and then click  **Developer Reference**). 
    
    For more details about using the  **Run** method to call Save as Web Page, see "Calling the **Run** method of the SaveAsWeb add-on" later in this topic.
    

## Save as Web Page command-line options

The format for command-line parameters is as follows: / _option_= _value_

For example, the following code sets the  **target** parameter: /target=c:\temp\mypage.htm

The following table lists the command-line options for the Save as Web Page command-line interface. The "Method/Property name" column lists the corresponding method or property in the object model. For details about a particular option, see the corresponding method or property topic in this reference.



|**Option**|**Default**|**Value type**|**Method/Property name**|
|:-----|:-----|:-----|:-----|
|target|None. You must supply a target value or Visio will generate an error.|Text| [TargetPath](viswebpagesettings-targetpath-property-visio-save-as-web.md)|
|pagetitle|Same as document file name|Text| [PageTitle](viswebpagesettings-pagetitle-property-visio-save-as-web.md)|
|prop|TRUE|Boolean| [PropControl](viswebpagesettings-propcontrol-property-visio-save-as-web.md)|
|altformat|TRUE|Boolean| [AltFormat](viswebpagesettings-altformat-property-visio-save-as-web.md)|
|folder|TRUE|Boolean| [StoreInFolder](viswebpagesettings-storeinfolder-property-visio-save-as-web.md)|
|theme|Null|Text| [ThemeName](viswebpagesettings-themename-property-visio-save-as-web.md)|
|startpage|-1 (all pages)|Number| [StartPage](viswebpagesettings-startpage-property-visio-save-as-web.md)|
|endpage|-1 (all pages)|Number| [EndPage](viswebpagesettings-endpage-property-visio-save-as-web.md)|
|openbrowser|TRUE|Boolean| [OpenBrowser](viswebpagesettings-openbrowser-property-visio-save-as-web.md)|
|screenres|1024x768|Text/Number1| [DispScreenRes](viswebpagesettings-dispscreenres-property-visio-save-as-web.md)|
|priformat|XAML|Text/Number1| [PriFormat](viswebpagesettings-priformat-property-visio-save-as-web.md)|
|secformat|PNG|Text/Number1| [SecFormat](viswebpagesettings-secformat-property-visio-save-as-web.md)|
|silent|FALSE|Boolean| [SilentMode](viswebpagesettings-silentmode-property-visio-save-as-web.md)|
|quiet|FALSE|Boolean| [QuietMode](viswebpagesettings-quietmode-property-visio-save-as-web.md)|
|stylesheet|\ _your_Visio_path\your_language_ID_\Default.css|Text| [Stylesheet](viswebpagesettings-stylesheet-property-visio-save-as-web.md)|
|navbar|TRUE|Boolean| [NavBar](viswebpagesettings-navbar-property-visio-save-as-web.md)|
|search|TRUE|Boolean| [Search](viswebpagesettings-search-property-visio-save-as-web.md)|
|panzoom|TRUE|Boolean| [PanAndZoom](viswebpagesettings-panandzoom-property-visio-save-as-web.md)|
1For the text/number value type, the user may specify text such as  _vml_ for the output type, or a number (for example, 1) representing the index of this output type. Each output type will have its own unique index. For **screenres**, text and number values are defined by the  [VISWEB_DISP_RES](visweb_disp_res-enumeration-visio-save-as-web.md) enumeration.


## Using the RUNADDONWARGS function

The following shows one way to use the RUNADDONWARGS function to call the SaveAsWeb add-on.


```
=RUNADDONWARGS("SaveAsWeb","/target=c:\temp\mypage.htm /quiet /prop /startpage=1 /endpage=3 /altformat /priformat=vml /secformat=jpg /openbrowser")
```

A scenario previously mentioned in this topic described a user being able to merely double-click a shape in a drawing to produce a Web page for that drawing. To demonstrate this, you can place the previous formula in the EventDblClick cell of the Events section in the ShapeSheet window of any shape on your drawing page. (To open the ShapeSheet window, select a shape in the drawing window, and then on the  **Developer** tab, click **Show ShapeSheet**.) After the formula is entered in the ShapeSheet cell, you can double-click that shape in the drawing window to launch the Save as Web Page feature.

For more information about the RUNADDONWARGS function, the EventDblClick cell, and the Events section, see the Visio Developer Reference (click  **Help**, click  **Search**, and then click  **Developer Reference**).


## Calling the Run method of the SaveAsWeb add-on

The Save as Web Page feature is installed as a Visio add-on called SaveAsWeb. To get a reference to this add-on, use the  **Addons** collection of the Visio **Application** object.

The following example shows how to run the SaveAsWeb add-on by passing command-line parameters to the  **Run** method of the **Addon** object.

In this example, the code that launches the add-on is contained in an event handler for the  **DocumentSaved** event. The **QuietMode** property is set to **True** so that the **Save as Web Page** dialog boxes are not displayed in the user interface.




```vb
Private Sub Document_DocumentSaved(ByVal Document As IVDocument) 
    Application.Addons("SaveAsWeb").Run "/quiet=True /target=C:\temp\test.htm" 
End Sub
```

For more information about the  **Addons** collection, the **Application** and **Addon** objects and the **DocumentSaved** event, see the Visio Automation Reference (click **Help**, click  **Search**, and then click  **Developer Reference**).


