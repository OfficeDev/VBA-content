---
title: About Programming Visio Viewer
ms.prod: visio
ms.assetid: b9bf841f-c5e5-c737-b8c0-86bd171144c8
ms.date: 06/08/2017
---


# About Programming Visio Viewer

Visio Viewer is an ActiveX control that lets you open, view, or print Visio drawings, even if you do not have Visio. You cannot, however, edit, save, or create a new Visio drawing in Visio Viewer. For that, you need Visio.

Visio Viewer provides an application programming interface (API) that lets solution developers perform numerous tasks, among them the following:

- Load and unload Visio drawings.
    
- Select shapes.
    
- Follow hyperlinks.
    
- Display various Visio Viewer dialog boxes to the user.
    
- Customize the size and position of the Visio Viewer window.
    
- Customize the user interface by changing foreground and background colors and displaying or hiding the grid and the scroll bars.
    
- Control the color and transparency of layers in the drawing.
    
- Control the color and visibility of reviewer markups (comments).
    
- Customize the toolbar by adding or removing buttons.
    
- Respond to user actions in the Visio Viewer interface.
    

## Programming Visio Viewer in Visual Basic 6.0

You can use Visual Basic 6.0 to instantiate the Visio Viewer control in various containers (for example, a Windows form). Before you can do so, you must get a reference to the Visio Viewer API.

Use the following steps to get a reference to the Visio Viewer API in a Visual Basic 6.0 project.


### To get a reference to the Visio Viewer API in a Visual Basic 6.0 project


1. Open Visual Basic 6.0. In Windows Vista or Windows 7, right-click the program shortcut, and then click  **Run as administrator**
    
2. In Visual Basic 6.0, open a new  **Standard EXE** project.
    
3. In your project, right-click the Toolbox, click  **Components**, select  **Microsoft Visio Viewer 14.0 Type Library**, and then click  **OK**.
    
4. Before you compile your code, on the  **Project** menu, click [ _your project name_]  **Properties**.
    
5. On the  **Make** tab, verify that **Remove information about unused ActiveX controls** is not selected.
    
The following code shows how to instantiate Visio Viewer in a form in Visual Basic 6.0. It creates a Visio Viewer control, displays the  **Properties and Settings** dialog box, sets the location, size, and visibility of the control within the form, and loads a document named "MyFile.vsd" into the control.

Add the following code to the project you created.




```vb
Dim Viewer1 As VisioViewerCtl.Viewer

Private Sub Form_Load()
    
    Set Viewer1 = Form1.Controls.Add("VisioViewer.Viewer", "Viewer1", Form1)

    Viewer1.Visible = True
    Viewer1.Left = 200
    Viewer1.SRC = "C:\Users\<variable>username</variable>\Documents\MyFile.vsd"

    Viewer1.Height = 5000
    Viewer1.Width = 5000
    Viewer1.DisplayPropertyDialog

End Sub
```


## Programming Visio Viewer on an HTML (Web) page

You can use the Visio Viewer control to embed a Visio drawing into a Web page, by manually inserting tags and parameters in the source code of the page. To write the source code, you can use a text editor, such as Notepad, or any other application that creates Web pages, such as Microsoft Expression Web 3 or SharePoint Designer.

You can set any of the properties of Visio Viewer by using the PARAM tag, as shown in the following sample code, which sets the  **Src** property of Visio Viewer.

Remember that because Visio Viewer is an ActiveX control, its behavior is influenced by Internet Explorer security settings.

The following code shows how to open a Visio drawing file in a Visio Viewer control on a Web page. It sets the height and width of the Visio Viewer control on the page and loads a source document into the control.

Copy the code into a file in a text editor, and save the resulting document as an HTM file. The Visio document "SalesData.vsd" referenced by the  **Src** parameter should be in the same folder as the HTM file.




```HTML
<html>
<OBJECT id="DrawingControl1"
    height=400 
    width=600
    classid="clsid:279D6C9A-652E-4833-BEFC-312CA8887857" VIEWASTEXT>
<PARAM NAME="Src" VALUE="SalesData.vsd">
</OBJECT>
</html>
```


## Programming Visio Viewer in managed code

You can use managed code to instantiate the Visio Viewer control in various containers, such as a Windows form, for example. Before you can do so, you must get a reference to the Visio Viewer API.

Use the following steps to get a reference to the Visio Viewer API in a Visual Studio 2008 project.


### To get a reference to the Visio Viewer API in a Visual Studio project


1. On the  **Start** menu, point to **All Programs**, click  **Accessories**, and then click  **Command Prompt** to open the **Command Prompt** window.
    
2. In the Command Prompt window, navigate to the Microsoft Office/Office 14 subfolder of the Program Files folder.
    
3. Copy the file VViewer.dll to a folder location to which you have permission to write new files (for example, your user folder).
    
4. Close the Command Prompt window, and then open the Visual Studio 2008 Command Prompt window. (On the  **Start** menu, point to **All Programs**, click  **Microsoft Visual Studio 2008**, click  **Visual Studio Tools**, and then click  **Visual Studio 2008 Command Prompt**).
    
5. In the Visual Studio 2008 Command Prompt window, navigate to the folder to which you copied the Visio Viewer DLL file.
    
6. In that folder, type AxImp.exe vviewer.dll to generate several files, including AxVisioViewer.dll.
    
7. In Visual Studio 2008, open a new Windows Forms Application project.
    
8. In your project, on the  **Project** menu, click **Add Reference**, and then click  **Browse**.
    
9. Browse to the folder where you created the AxVisioViewer.dll file, select that file in the list, and then click  **OK**.
    
In your Visual Studio project, in the Form1.cs file, add the following code to instantiate the Visio Viewer control, set some of its properties, and load a test file into the control. This code assumes that you have a Visio file named "test.vsd" in your Documents folder, at the path shown. Modify the path and file names accordingly for your computer.




```
<code language="CS" type="developer">public partial class Form1 : Form
    {
        private AxVisioViewer.AxViewer viewer;

        /// &;lt;summary&;gt;
        /// The Visio Viewer OM
        /// &;lt;/summary&;gt;
        public AxVisioViewer.AxViewer Viewer
        {
            get
            {
                return this.viewer;
            }
        }

        public Form1()
        {
            this.InitializeComponent();
            this.Resize += new EventHandler(this.UpdateSize);
            this.viewer = new AxVisioViewer.AxViewer();
            this.Controls.Add(this.viewer);
            this.viewer.CreateControl();

            this.viewer.Location = new Point(0, 0);
            this.UpdateSize(null, null);
         
        }

        public void UpdateSize(object obj, EventArgs ea)
        {
            this.viewer.ClientSize = new Size(this.ClientSize.Width - 150, this.ClientSize.Height - 150);
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            this.viewer.Load("C:\\users\\username\\documents\\viewer\\test.vsd");

        }        

    }
</code>
```


