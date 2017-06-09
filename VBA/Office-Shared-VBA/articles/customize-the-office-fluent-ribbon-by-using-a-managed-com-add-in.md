---
title: Customize the Office Fluent Ribbon by Using a Managed COM Add-in
ms.prod: office
ms.assetid: 7926e6bc-c7ae-cc6f-faa5-28e2e6de664c
ms.date: 06/08/2017
---


# Customize the Office Fluent Ribbon by Using a Managed COM Add-in

The ribbon component of the Microsoft Office Fluent user interface in Microsoft Office suites gives users a flexible way to work with Office applications. Ribbon Extensibility (RibbonX) uses simple, text-based, declarative XML markup to create and customize the ribbon.

The code example in this topic shows how to customize the ribbon in an Office application, no matter what document is open. In the following steps, you create application-level customizations by using a managed COM add-in, and you create the add-in in Microsoft Visual Studio 2012 by using Microsoft Visual C#. The project adds a custom tab, a custom group, and a custom button to the ribbon. To complete the procedure, you perform the following tasks.

1. Create the XML customization file.
    
2. Create a managed COM add-in project in Microsoft Visual Studio 2012 with C#.
    
3. Add the XML customization file to the project as an embedded resource.
    
4. Implement the  **IRibbonExtensibility** interface.
    
5. Create a callback method that is triggered when the button is clicked.
    
6. Build, install, and test the project.
    
 **Create the XML Customization File**
In this step, you create the file that adds the custom components to the ribbon. 

1. In a text editor, add the following XML markup. 
    
  ```XML
  <customUI xmlns="http://schemas.microsoft.com/office/2006/01/customui"> 
  <ribbon> 
    <tabs> 
      <tab id="CustomTab" label="My Tab"> 
        <group id="SampleGroup" label="Sample Group"> 
          <button id="Button" label="Insert Company Name" size="large" onAction="InsertCompanyName" /> 
        </group > 
      </tab> 
    </tabs> 
  </ribbon> 
</customUI> 

  ```

2. Close and save the file as  **customUI.xml**.
    

 **Create the Managed COM Add-in Project**
In this step, you create a COM add-in C# project in Microsoft Visual Studio 2012.

1. Start Microsoft Visual Studio 2012.
    
2. On the  **File** menu, click **New Project**.
    
3. In the  **New Project** dialog box under **Project Types**, expand  **Other Projects**, click  **Extensibility Projects**, and then double-click  **Shared Addin**.
    
4. Add a name for the project; for this sample, type  **RibbonXSampleCS**.
    
5. In the first screen of the  **Shared Add-in Wizard**, click  **Next**.
    
6. Select  **Create an Add-in using Visual C#**, and then click  **Next**.
    
7. Clear all of the selections except  **Microsoft Word** and then click **Next**.
    
8. Type a name and description for the add-in, and then click  **Next**.
    
9. In the  **Choose Add-in Options** screen, select **I would like my Add-in to load when the host application loads** and then click **Next**.
    
10. Click  **Finish** to complete the wizard.
    

 **Add External References to the Project**
In this step you add references to the Word Primary Interop Assemblies and type library. 

1. In the Solution Explorer, right-click  **References**, and then click  **Add Reference**.
    
     **Note**  If you do not see the  **References** folder, click the **Project** menu and then click **Show All Files**.
2. Scroll down on the  **.NET** tab, press the **CTRL** key, and then select **Microsoft.Office.Interop.Word**.
    
3. On the  **COM** tab, scroll down, select either the **Microsoft Office 15.0 Object Library** (or the library that is appropriate for your version of Office), and then click **OK**.
    
4. Add the following namespace references to the project, if they do not already exist, just below the  **namespace** line.
    
  ```C#
  using System.Reflection; 
using Microsoft.Office.Core; 
using System.IO; 
using System.Xml; 
using Extensibility; 
using System.Runtime.InteropServices; 
using MSword = Microsoft.Office.Interop.Word; 

  ```


 **Add the Customization File as an Embedded Resource**
In this step, you add the XML customization file as an embedded resource in the project.

1. In the Solution Explorer, right-click  **RibbonXSampleCS**, point to  **Add**, and click  **Existing Item**.
    
2. Navigate to the  **customUI.xml** file that you created, select the file, and then click **Add**.
    
3. In the Solution Explorer, right-click  **customUI.xml**, and then select  **Properties**.
    
4. In the  **Property** window select **Build Action**, and then scroll down to  **Embedded Resource**.
    

 **Implement the IRibbonExtensibility Interface**
In this step you add code to the Extensibility.IDTExtensibility2::OnConnection to create a reference to the Word application at runtime. You also implement the only member of the  **IRibbonExtensibility** interface, **GetCustomUI**.

1. In the Solution Explorer, right-click  **Connect.cs**, and click  **View Code**.
    
2. After the  **Connect** method, add the following declaration, which creates a reference to the **Word Application** object:
    
     `private MSword.Application applicationObject;`
    
3. Add the following line to the  **OnConnection** method. This statement creates an instance of the **Word Application** object:
    
     `applicationObject =(MSword.Application)application;`
    
4. At the end of the public class Connect statement, add a comma, and then type  **IRibbonExtensibility**.
    
     **Note**  You can use Microsoft IntelliSense to insert interface methods for you. For example, at the end of the public class Connect: statement, type  **IRibbonExtensibility**, right-click and point to **Implement Interface**, and then click  **Implement Interface Explicitly**. This adds a stub for the  **GetCustomUI** method. The implementation looks similar to the following code.

  ```C#
  string IRibbonExtensibility.GetCustomUI(string RibbonID) 
{ 
}
  ```

5. Insert the following statement into the  **GetCustomUI** method, overwriting the existing code. `return GetResource("customUI.xml");`
    
6. Insert the following method below the  **GetCustommUI** method:
    
  ```C#
  private string GetResource(string resourceName) 
        { 
            Assembly asm = Assembly.GetExecutingAssembly(); 
            foreach (string name in asm.GetManifestResourceNames()) 
            { 
                if (name.EndsWith(resourceName)) 
                { 
                    System.IO.TextReader tr = new System.IO.StreamReader(asm.GetManifestResourceStream(name)); 
                    //Debug.Assert(tr != null); 
                    string resource = tr.ReadToEnd(); 
 
                    tr.Close(); 
                    return resource; 
                } 
            } 
            return null; 
        } 

  ```


    The  **GetCustomUI** method calls the **GetResource** method. The **GetResource** method sets a reference to this assembly during runtime and then loops through the embedded resource until it finds the one named customUI.xml. It then creates an instance of the **StreamReader** object that reads the embedded file containing the XML markup. The procedure passes the XML back to the **GetCustomUI** method which returns the XML to the ribbon. Alternately, you can construct a string that contains the XML markup and read it directly into the **GetCustomUI** method.
    
7. Following the  **GetResource** method, add this method. This method inserts the company name into the document at the beginning of the page.
    
  ```C#
  public void InsertCompanyName(IRibbonControl control) 
        { 
        // Inserts the specified text at the beginning of a range or selection. 
            string MyText; 
            MyText = "Microsoft Corporation"; 
 
            MSword.Document doc = applicationObject.ActiveDocument; 
 
            //Inserts text at the beginning of the active document. 
            object startPosition = 0; 
            object endPosition = 0; 
            MSword.Range r = (MSword.Range)doc.Range( 
                   ref startPosition, ref endPosition); 
            r.InsertAfter(MyText); 
        } 

  ```


 **Build and Install the Project**
In this step, you build the add-in and its setup project. Before you continue, make sure that Word is closed.

1. In the  **Project** menu, click **Build Solution**. When the build is complete, a notification appears in the lower left corner of the window.
    
2. In the Solution Explorer, right-click  **RibbonXSampleCSSetup** and click **Build**.
    
3. Right-click  **RibbonXSampleCSSetup** again and click **Install** to begin the **RibbonXSampleCSSetup Setup Wizard**.
    
4. Click  **Next** in each of the screens and then click **Close** on the final screen.
    
5. Start Word. You should see the  **My Tab** tab to the right of the other tabs.
    

 **Test the Project**
Click the  **My Tab** tab and then click **Insert Company Name** to insert the company name into the document at the cursor. If you do not see the customized ribbon, you might need to add an entry to the Windows registry by completing the following steps.

 **Caution**  The next few steps contain information about how to modify the registry. Before you modify the registry, be sure to back it up and make sure that you understand how to restore the registry if a problem occurs. For more information about how to back up, restore, and edit the registry, search for the following article in the Microsoft Knowledge Base:  **256986 Description of the Microsoft Windows Registry**.


1. In the Solution Explorer, right-click the setup project,  **RibbonXSampleCSSetup**, point to  **View**, and then click  **Registry**.
    
2. From the  **Registry** tab, navigate to the following registry key for the add-in: HKCU\Software\Microsoft\Office\Word\AddIns\RibbonXSampleCS.Connect
    
     **Note**  If the  **RibbonXSampleCS.Connect** key does not exist, you can create it. To do so, right-click the **Addins** folder, point to **New**, and then click  **Key**. Name the key  **RibbonXSampleCS.Connect**. Add a  **LoadBehavior** **DWord** and set its value to **3**.


