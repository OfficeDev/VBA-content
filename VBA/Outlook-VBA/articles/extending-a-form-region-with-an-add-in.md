---
title: Extending a Form Region with an Add-in
ms.prod: outlook
ms.assetid: b1a28a20-a0b8-cc57-7672-da51ec8bb097
ms.date: 06/08/2017
---


# Extending a Form Region with an Add-in



While you can create and run forms with form regions without a COM add-in, using a COM add-in will allow form regions to include custom business logic or advanced functionality. Unlike customizing form pages in a standard form, you do not use VBScript to write code behind a form; instead, you program form regions with a COM add-in. Your add-in will implement a new interface,  **[FormRegionStartup](formregionstartup-object-outlook.md)**. Add-ins will be able to use Microsoft Forms 2.0 controls and Microsoft Outlook controls in a form region. This topic describes how to implement  **FormRegionStartup** and access Outlook controls in a form region.

## Specifying the Use of an Add-in

When you register the form region for a message class, create a key in the Windows registry for that message class (if the key does not yet exist), and specify as data, an equal sign ( **=**) followed by the ProgID of the add-in. For more information on registering a form region in the Windows registry, see  [Specifying Form Regions in the Windows Registry](specifying-form-regions-in-the-windows-registry.md).


## Implementing FormRegionStartup

In the same class that implements the  **IDTExtensibility2** interface of your COM Add-in, implement the **Outlook.FormRegionStartup** interface, which is defined in the Outlook type library. Outlook will call the four methods in this interface:


-  **[GetFormRegionStorage](formregionstartup-getformregionstorage-method-outlook.md)**
    
-  **[BeforeFormRegionShow](formregionstartup-beforeformregionshow-method-outlook.md)**
    
-  **[GetFormRegionManifest](formregionstartup-getformregionmanifest-method-outlook.md)**
    
-  **[GetFormRegionIcon](formregionstartup-getformregionicon-method-outlook.md)**
    



## GetFormRegionStorage

When Outlook is about to display a form region that is controlled by an add-in, Outlook will call the  **GetFormRegionStorage** method. When the add-in receives a call from Outlook to **GetFormRegionStorage** specifying information for a form region, the add-in will return information for the layout. This information can be a local path to the layout file (.OFS file), a Microsoft Windows **IStorage** object, or a byte array with the contents of the OFS file, which allows an add-in to store the OFS as a resoruce. Outlook will use the returned information to instantiate controls and calculate the layout for the form region. Outlook will also instantiate a **[FormRegion](formregion-object-outlook.md)** object for the form region. The method prototype for **GetFormRegionStorage** in Microsoft Visual Basic and Microsoft C# are shown below.

In Visual Basic: 




```VB.net
Public Function GetFormRegionStorage(ByVal FormRegionName As String,
    ByVal Item As Object, 
    ByVal LCID As Integer, 
    _ ByVal FormRegionMode As Outlook.OlFormRegionMode, 
    ByVal FormRegionSize As Outlook.OlFormRegionSize) _ 
Implements Microsoft.Office.Interop.Outlook.FormRegionStartup.GetFormRegionStorage 

```

In C#: 




```C#
public object GetFormRegionStorage(string FormRegionName, 
    object Item, 
    int LCID,
    Outlook.OlFormRegionMode FormRegionMode, 
    Outlook.OlFormRegionSize FormRegionSize) 
```


## BeforeFormRegionShow

If  **GetFormRegionStorage** is successful, just before the form region is displayed in an Inspector window or in the Reading Pane, Outlook will call **BeforeFormRegionShow**, passing the  **FormRegion** object to the add-in. The add-in will use this chance before the form region is displayed to update anything in the user interface, such as changing label captions, as described in the section **Accessing Outlook Controls** below, and suppressing irrelevant content. The Visual Basic and C# method prototypes for **BeforeFormRegionShow** are shown below.

In Visual Basic: 




```VB.net
Public Sub BeforeFormRegionShow(ByVal Item As Object, 
    ByVal FormRegion As Microsoft.Office.Interop.Outlook.FormRegion) _ 
Implements Microsoft.Office.Interop.Outlook.FormRegionStartup.BeforeFormRegionShow 

```

In C#: 




```C#
public void BeforeFormRegionShow(object Item, Outlook.FormRegion FormRegion) 
```


## Accessing Outlook Controls

When using a COM add-in to extend a form region, you are often listening to control events, calling control methods, or reading and setting control properties. To access Microsoft Forms 2.0 controls, Outlook controls, or the form canvas object in an add-in, you must add a reference to the Microsoft Forms 2.0 object library. Adding this reference will give you access to the  **Microsoft.Vbe.Interop.Forms** namespace in your add-in project.

After adding the reference, optionally, you can create an alias to the namespace of the type library to make it easier to use the included types. To create an alias, insert the following code at the top of the code file. The following are examples of how to do this if you are writing your add-in in Visual Basic or C#. These aliases will also be used in the code samples further below.

In Visual Basic: 




```VB.net
Imports Outlook = Microsoft.Office.Interop.Outlook 
Imports Office = Microsoft.Office.Core  
Imports MSForms = Microsoft.Vbe.Interop.Forms 
```

In C#: 




```C#
using Outlook = Microsoft.Office.Interop.Outlook; 
using Office = Microsoft.Office.Core; 
using MSForms = Microsoft.Vbe.Interop.Forms; 

```

You can access controls through the  **FormRegion** object obtained from **BeforeFormRegionShow**. The  ** [FormRegion.Form](formregion-form-property-outlook.md)** property returns an object representing a form; you can cast this object to the **MSForms.UserForm** class (exposed in the Microsoft Forms 2.0 object library) to access the form canvas for the form region.

Each instance of the  **UserForm** object has a **Controls** collection that can be used to access the individual controls on the **UserForm** by control name. Many of the Microsoft Forms 2.0 controls have themed counterparts that are Outlook controls. In a form region, Outlook replaces those Forms 2.0 controls that have Outlook counterpart controls by the corresponding themed counterparts. Once you have obtained a reference to a themed control from the **Controls** collection, you can cast it to the proper type in the Outlook type library. You will then be able to access all the properties, methods, and events exposed for these controls in the Outlook type library. Unlike customizing forms with VBScript, you will be able to listen to all control events, and not only the **Click** event. For more information on controls, see [Controls in a Custom Form](controls-in-a-custom-form.md).

The following code samples show how the  **BeforeFormRegionShow** method uses the input parameterFormRegion from Outlook to obtain a form object, then casts it to an **MSForms.UserForm** class and accesses the collection of controls in the **UserForm** object. The form canvas represented by this **UserForm** object has two Outlook controls: a text box named `OlkTextBox1` and a check box named `OlkCheckBox1`. It casts them to the proper Outlook control types and sets default values for these controls as follows. 

In Visual Basic:




```VB.net
Dim UserForm As MSForms.UserForm 
Dim FormControls As MSForms.Controls 
Dim TextBox1 As Outlook.OlkTextBox 
Dim CheckBox1 As Outlook.OlkCheckBox 
 
UserForm = FormRegion.Form 
FormControls = UserForm.Controls 
 
TextBox1 = FormControls.Item("OlkTextBox1") 
TextBox1.Text = "Sample Form Region" 
CheckBox1 = FormControls.Item("OlkCheckBox1") 
CheckBox1.Value = True 

```

In C#:




```C#
MSForms.UserForm userForm = (MSForms.UserForm)FormRegion.Form; 
MSForms.Controls formControls = userForm.Controls; 
 
Outlook.OlkTextBox textBox1 =  
   (Outlook.OlkTextBox)formControls.Item("OlkTextBox1"); 
textBox1.Text = "Sample Form Region"; 
 
Outlook.OlkCheckBox checkBox1 =  
   (Outlook.OlkCheckBox)formControls.Item("OlkCheckBox1"); 
checkBox1.Value = true; 

```


## GetFormRegionManifest

When Outlook starts, it reads the list of form regions from the Windows registry and caches the data. Based on this data, if Outlook notices that an add-in is to provide the XML manifest for a form region, Outlook will use the ProgID provided in the cached data and call the  **GetFormRegionManifest** method implemented by this add-in to obtain the XML it needs to display the form region. If the XML manifest is not valid and does not conform to the form region XML schema, Outlook will not be able to load the form region.

For more information on specifying a ProgID when registering a form region, see  [Specifying Form Regions in the Windows Registry](specifying-form-regions-in-the-windows-registry.md).

The Visual Basic and C# method prototypes for  **GetFormRegionManifest** are shown below.

In Visual Basic: 




```VB.net
Public Function GetFormRegionManifest(ByVal FormRegionName As String, 
    ByVal LCID As Integer) _ 
Implements Microsoft.Office.Interop.Outlook.FormRegionStartup.GetFormRegionManifest 

```

In C#: 




```C#
public object GetFormRegionManifest(string FormRegionName, int LCID)
```


## GetFormRegionIcon

When Outlook starts, it reads the list of form regions from the Windows registry and caches the data associated with the form regions. If a form region has been registered with a ProgID, Outlook will resort to the corresponding add-in by calling its implementation of  **GetFormRegionIcon** for any icon in the XML manifest that has `addin` as the value of a child element of the **icons** element. For more information on using an add-in to specify icons, see [How to: Use an Add-in to Specify Icons for a Form Region](use-an-add-in-to-specify-icons-for-a-form-region.md).

The Visual Basic and C# method prototypes for  **GetFormRegionIcon** are shown below.

In Visual Basic: 




```VB.net
Public Function GetFormRegionIcon(ByVal FormRegionName As String, 
    ByVal LCID As Integer, _ 
    ByVal Icon As Outlook.OlFormRegionIcon) _ 
Implements Microsoft.Office.Interop.Outlook.FormRegionStartup.GetFormRegionManifest 

```

In C#: 




```C#
public object GetFormRegionIcon(string FormRegionName, int LCID, Outlook.OlFormRegionIcon Icon)
```


