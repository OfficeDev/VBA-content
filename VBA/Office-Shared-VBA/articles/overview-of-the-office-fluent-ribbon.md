---
title: Overview of the Office Fluent Ribbon
ms.prod: office
ms.assetid: 773c202c-f5f9-c4f6-f833-0dd56eb21a8f
ms.date: 06/08/2017
---


# Overview of the Office Fluent Ribbon

## 


 **Note**  The use of CommandBars in some Microsoft Office applications has been superseded by the new ribbon component of the Microsoft Office Fluent user interface. For more information, search Help for the keyword "ribbon."

The Office Fluent ribbon replaces the previous system of layered menus, toolbars, and task panes from previous versions of Office. The ribbon has a simpler system of interfaces that is optimized for efficiency and discoverability. The ribbon has improved context menus, screentips, a mini toolbar, and keyboard shortcuts that improve user efficiency and productivity. In addition, you can use Ribbon Extensibility (or RibbonX) to enhance the user experience. You use extensible markup language (XML) and one of several conventional programming languages to manipulate the components that make up the ribbon. Because XML is plain text, you can create customization files in any text editor, or you can use your favorite XML editor. You can also reuse customization files with a minimum of adjustments because each application uses the same programming model. For example, you can reuse the customization files that you create in Word, Excel, Access, or PowerPoint.
Using XML markup files to customize the ribbon greatly reduces the need to create complex add-ins that are based on the  **CommandBars** object model. However, add-ins written in previous versions of Office continue to work in the ribbon with little or no modification. You can create application-level customizations for the ribbon in Word, in Excel, or PowerPoint by doing any of the following:

- Use COM add-ins in managed or unmanaged code
    
- Use application-specific add-ins, such as .ppam and .xlam files
    
- Use templates (.dotm files) in Word
    

In a typical scenario, code in the COM add-in contains procedures that return XML markup from an external customization file or from XML that is in the code itself. When the application starts, the add-in loads and executes the code that returns the XML markup. Microsoft Office validates the XML markup against an XSD schema, and then loads it into memory and applies it to the ribbon before the ribbon is displayed. Menu items and controls use callback procedures to execute code in the add-in. Document-level customizations use the same XML markup and an Open XML Formats file with one of these extensions: docx, .docm, .xlsx, .xlsm, .pptx, .or pptm. In this scenario, you create a customization file that contains the XML markup and save it to a folder. You then modify the parts in the Open XML Formats container to point to the customization file. When you open the document in the Office application, the customization file loads into memory and is applied to the ribbon. The commands and controls then call code that is contained in the document to provide their functionality.
 **What About Existing Solutions?**
In versions of Microsoft Office previous to Office 2007, developers used the  **CommandBars** object model to create the Microsoft Visual Basic® code that modified the UI. In Office, this legacy code continues to work in most cases without modification. However, changes made to toolbars in Office 2003 now appear on an **Add-Ins** tab in the Office. The type of customization that appears depends on the original design of the add-in. For example, Office creates a **Menu Commands** group that contains the items that you added to the previous menu structure ( **File** menu, **Insert** menu, **Tools** menu, and so forth). It also creates a **Toolbar Commands** group that contains the items that you added to the previous built-in toolbars (such as the **Standard** toolbar, **Formatting** toolbar, and **Picture** toolbar). In addition, custom toolbars from an add-in or a document appear in the **Custom Toolbars** group on the **Add-Ins** tab.
 **Callback Procedures Add Functionality to the Ribbon**
With Ribbon Extensibility, you specify callbacks to update properties and perform actions from your UI at runtime. For example, consider the  **onAction** callback method for a button in the following RibbonX markup. `<button id="myButton" onAction="MyButtonOnAction" />` This markup tells Office to call the MyButtonOnAction function when the button is clicked. The MyButtonOnAction function has a specific signature that depends on your choice of languages; here is an example in Microsoft Visual C#.



```C#
public void MyButtonOnAction (IRibbonControl control) 
   { 
      if (control.Id=="myButton") 
      { 
         System.Windows.Forms.MessageBox.Show("Button clicked!"); 
      } 
   } 
```

 **Customizing the Ribbon with COM Add-Ins**
Customization at the application-level results in a modified ribbon that appears in the application no matter which document is open. You create COM add-ins primarily to make these modifications. To customize the ribbon by using COM add-ins, do the following: 

1. Create a COM add-in project. The add-in that you create must implement the Extensibility.IDTExtensibility2 interface that all COM add-ins implement, as well as the  **IRibbonExtensibility** interface that is in the Microsoft.Office.Core namespace.
    
2. Build the add-in and setup project, and then install the project.
    
3. Start the Office application. When the add-in loads, it triggers the IDTExtensibility2::OnConnection event, which initializes the add-in, just as in previous versions of Office.
    
4. Next, the  **QueryInterface** method is called, which determines if the **IRibbonExtensibility** interface is implemented.
    
5. If so, the  **IRibbonExtensibility::GetCustomUI** method is called, which loads the XML markup from the XML customization file or from XML markup embedded in the procedure, and then loads the customizations into the application.
    
6. The customized UI is now ready for the user.
    

 **Customizing the Ribbon with Office Open XML Formats Files**
To customize the UI by using XML markup, do the following: 

1. Create the customization file in any text editor. Add XML markup that adds new components to the ribbon, modifies existing components, or hides components. Save the file as  **customUI.xml**.
    
2. Create a folder on your desktop named  **customUI** and copy the customization file to the folder.
    
3. Validate the XML markup with custom UI schema. 
    
     **Note**  This step is optional.
4. Create a document in the Office application, and then save it as an Open XML Formats file with one of these extensions:  _.docx_, _.docm_, _.xlsx_, _.xlsm_, _.pptm_, or _.pptx_. For security, files that contain macros have an _m_ suffix, and can contain procedures that are called by RibbonX commands and controls.
    
5. Add a  _.zip_ extension to the document file name, and then open the file.
    
6. Add the customization file to the container by dragging the folder to the file.
    
7. Extract the  **.rels** file that is in the .zip file to your desktop. A **_rels** folder that contains the .rels file is copied to your desktop.
    
8. Open the .rels file and add a line that creates a relationship between the document file and the customization file, and then save the file.
    
9. Add the _rels folder back to the container, overwriting the existing file.
    
10. Rename the file to its original name by removing the .zip extension. When you open the Office file, the ribbon appears, including your customization to the ribbon.
    

 **General Format of XML Markup Files**
You can use XML markup to customize the ribbon. The following example shows the general format of an XML markup file that you can use to customize the ribbon in Word.



```XML
<customUI xmlns="http://schemas.microsoft.com/office/2006/01/customui"> 
  <ribbon> 
    <tabs> 
      <tab idMso="TabHome"> 
        <group idMso="GroupFont" visible="false" /> 
      </tab> 
      <tab id="CustomTab" label="My Tab"> 
        <group id="SampleGroup" label="Sample Group"> 
          <toggleButton id="ToggleButton1" size="large" label="Large Toggle Button" getPressed="MyToggleMacro"  /> 
          <checkBox id="CheckBox1" label="A CheckBox" screentip="This is a check box" onAction="MyCheckboxMacro" /> 
          <editBox id="EditBox1" getText="MyTextMacro" label="My EditBox" onChange="MyEditBoxMacro"/> 
          <comboBox id="Combo1" label="My ComboBox" onChange="MyComboBoxMacro"> 
            <item id="Zip1" label="33455" /> 
            <item id="Zip2" label="81611" /> 
            <item id="Zip3" label="31561" /> 
          </comboBox> 
          <advanced> 
            <button id="Launcher1" screentip="My Launcher" onAction="MyLauncherMacro" /> 
          </advanced> 
        </group> 
        <group id="MyGroup" label="My Group" > 
          <button id="Button" label="My Large Button" size="large" onAction="MyButtonMacro" /> 
          <button id="Button2" label="My Normal Button" size="normal" onAction="MyOtherButtonMacro" /> 
        </group > 
      </tab> 
    </tabs> 
  </ribbon> 
</customUI> 

```

This sample makes the following changes to the ribbon in Word, in the following order:

1. Declares the default namespace and a custom namespace.
    
2. Hides the built-in  **GroupFont** group that is located on the built-in **Home** tab.
    
3. Adds a new  **CustomTab** tab to the right of the last built-in tab.
    
     **Note**  Use the  _id= identifier_ attribute to create a custom item, such as a custom tab. Use the _idMso= identifier_ attribute to refer to a built-in item, such as the **TabHome** tab.
4. Adds a new  **SampleGroup** group to the **My Tab** tab.
    
5. Adds a large-sized ToogleButton1 button to  **My Group** and specifies an onAction callback along with a GetPressed callback.
    
6. Adds a CheckBox1 check box to  **My Group** with a custom screentip and specifies an onAction callback.
    
7. Adds a EditBox1 edit box to  **My Group** and specifies an onChange callback.
    
8. Adds a Combo1 combo box to  **My Group** with three items. The combo box specifies an onChange callback that uses the text from each item.
    
9. Adds a Launcher1 launcher to  **My Group** with the onAction callback set. A launcher can also display a custom dialog box to offer more options to the user.
    
10. Adds a new  **MyGroup** group to the custom tab.
    
11. Adds a large-sized Button1 button to  **MyGroup** and specifies an onAction callback.
    
12. Adds a normal-sized Button1 button that is a normal-sized button to  **MyGroup** and specifies an onAction callback.
    

 **Working with Legacy Command Bar Add-Ins**
When you create COM add-ins, you usually need a way for users to interact with the add-in. In earlier versions of Office, you added a menu item or toolbar button to the application by using the  **CommandBars** object model. In this release of Office, custom applications continue to work in the ribbon without modification in most cases. However, changes that you made with the **CommandBars** object model, or any other technology that modified the menus or toolbars like WordBasic or XLM, appear on a separate **Add-Ins** tab. This makes it easier for users to locate the controls.
 **Dynamically Updating the Ribbon**
Callbacks that return properties of a control normally get called one time unless you specify that the call is to be repeated. You can requery your callback by implementing the onLoad callback in the CustomUI element. This callback gets called one time when the RibbonX markup file is successfully loaded, and then passes the code to an IRibbonUI object. The following code example gets the IRibbonUI object so that you can update your controls at runtime. 
XML Markup:
 `<customUI xmlns="http://schemas.microsoft.com/office/2006/01/customui" onLoad="ribbonLoaded">`
In C#: Write a callback in your Connect class.



```C#
IRibbonUI myRibbon; 
 
     public void ribbonLoaded(IRibbonUI ribbon) { 
         myRibbon = ribbon; 
     } 
```

The ribbon gives users a flexible way to work with Office applications. You use simple, text-based, declarative XML markup to create and customize the ribbon. With a few lines of XML, you can create just the right interface for the user. Because the XML markup is in one file, it is easier to modify the interface as requirements change. You can also improve user productivity by placing the commands where users can easily find them. Finally, the ribbon adds consistency across applications, which reduces the time that users spend to learn each application.

