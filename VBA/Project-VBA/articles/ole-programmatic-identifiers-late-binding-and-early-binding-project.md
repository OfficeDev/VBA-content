---
title: OLE programmatic identifiers, late binding, and early binding (Project)
keywords: vbapj.chm45320154
f1_keywords:
- vbapj.chm45320154
ms.prod: project-server
ms.assetid: c72f3f22-3628-1379-8c6b-79c7984c728d
ms.date: 06/08/2017
---


# OLE programmatic identifiers, late binding, and early binding (Project)

Learn how to add an  **Automation** object by using late binding at run time, and how to set a reference for early binding at design time.


## 

You can use an OLE programmatic identifier (sometimes called a ProgID) to create an automation object for run time binding. For example, if both Project and Word are installed on the computer, the following macro in Project creates a Word document named Doc1.docx, and then opens the  **Save As** dialog box in Word.


```vb
Sub CreateWordDoc_Late() 
    Dim wdDoc As Object 
 
    Set wdDoc = CreateObject("Word.Document") 
    wdDoc.Save 
End Sub
```


 **Note**  Objects created by using the ProgID have late binding at run time; therefore, you cannot see the object members available when you are writing code in the VBE. Late-bound objects also have poorer performance than objects created with early binding at design time. 

The following macro performs better and does the same job as the  **CreateWordDoc_Late** macro. The **CreateWordDoc_Early** macro requires that you add a reference to the **Microsoft Word 15.0 Object Library**. In the  **Tools** menu, choose **References** to open the **References - VBA Project** dialog box.




```vb
Sub CreateWordDoc_Early() 
    Dim wdDoc As Word.Document 
 
    Set wdDoc = New Word.Document 
    wdDoc.Save 
End Sub
```

Following is an example of using early binding to create an Excel worksheet. Set a reference to  **Microsoft Excel 15.0 Object Library**.




```vb
Sub CreateExcelWorkbook_Early()
    Dim xlApp As Excel.Application
    Dim xlWorkbook As Excel.Workbook
    Dim xlWorksheet As Excel.Worksheet
    
    Set xlApp = Excel.Application
    xlApp.Visible = True
        
    Set xlWorkbook = xlApp.Workbooks.Add
    Set xlWorksheet = xlWorkbook.Worksheets(1)
    
    xlWorksheet.Cells(1, 1).Value = "Data from Project"
    xlWorksheet.SaveAs ("C:\Project\VBA\ProjectWorksheet.xlsx")
    xlWorkbook.Close
    xlApp.Visible = False
End Sub
```

For information about using Project from another application, late binding, and early binding, see the [Application](application-object-project.md) object.

The following tables list OLE programmatic identifiers for ActiveX controls and several Microsoft Office applications.


 **Note**  Instead of using the ProgId values for late binding, we recommend that you set a reference to the equivalent object library and use early binding.

 **ActiveX Controls**

To create the ActiveX controls listed in the following table, use the corresponding OLE programmatic identifier. When you insert a user form, Project sets a reference to  **Microsoft Forms 2.0 Object Library** for early binding.



|**To create this control**|**Use this identifier**|
|:-----|:-----|
|CheckBox|Forms.CheckBox.1|
|ComboBox|Forms.ComboBox.1|
|CommandButton|Forms.CommandButton.1|
|Frame|Forms.Frame.1|
|Image|Forms.Image.1|
|Label|Forms.Label.1|
|ListBox|Forms.ListBox.1|
|MultiPage|Forms.MultiPage.1|
|OptionButton|Forms.OptionButton.1|
|ScrollBar|Forms.ScrollBar.1|
|SpinButton|Forms.SpinButton.1|
|TabStrip|Forms.TabStrip.1|
|TextBox|Forms.TextBox.1|
|ToggleButton|Forms.ToggleButton.1|
 **Microsoft Access**

To create the Access objects listed in the following table, use one of the corresponding OLE programmatic identifiers. If you use an identifier without a version number suffix, you create an object in the most recent version of Access that is available on the computer where the macro is running. For early binding, set a reference to  **Microsoft Access 15.0 Object Library**.



|**To create this object**|**Use one of these identifiers**|
|:-----|:-----|
|Application|Access.Application, Access.Application.15|
|CurrentData|Access.CodeData, Access.CurrentData|
|CurrentProject|Access.CodeProject, Access.CurrentProject|
|DefaultWebOptions|Access.DefaultWebOptions|
 **Microsoft Excel**

To create the Excel objects listed in the following table, use one of the corresponding OLE programmatic identifiers. If you use an identifier without a version number suffix, you create an object in the most recent version of Excel that is available on the computer where the macro is running. For early binding, set a reference to  **Microsoft Excel 15.0 Object Library**.



|**To create this object**|**Use one of these identifiers**|**Comments**|
|:-----|:-----|:-----|
|Application|Excel.Application, Excel.Application.15||
|Workbook|Excel.AddIn||
|Workbook|Excel.Chart, Excel.Chart.8|Returns a workbook containing two worksheets: one for the chart, and one for its data. The chart worksheet is the active worksheet.|
|Workbook|Excel.Sheet, Excel.Sheet.12|Returns a workbook with one worksheet.|
 **Microsoft Graph**

To create the Microsoft Graph objects listed in the following table, use one of the corresponding OLE programmatic identifiers. If you use an identifier without a version number suffix, you create an object in the most recent version of Graph that is available on the computer where the macro is running. For early binding, set a reference to  **Microsoft Graph 15.0 Object Library**.



|**To create this object**|**Use one of these identifiers**|
|:-----|:-----|
|Application|MSGraph.Application, MSGraph.Application.8|
|Chart|MSGraph.Chart, MSGraph.Chart.8|
 **Microsoft Office Web Components**


 **Note**  The Microsoft Office Web Component (OWC) is deprecated and is not installed with Project.

 **Microsoft Outlook**

To create the Microsoft Outlook object given in the following table, use one of the corresponding OLE programmatic identifiers. If you use an identifier without a version number suffix, you create an object in the most recent version of Outlook that is available on the computer where the macro is running. For early binding, set a reference to  **Microsoft Outlook 15.0 Object Library**.



|**To create this object**|**Use one of these identifiers**|
|:-----|:-----|
|Application|Outlook.Application, Outlook.Application.15|
 **Microsoft PowerPoint**

To create the Microsoft PowerPoint object given in the following table, use one of the corresponding OLE programmatic identifiers. If you use an identifier without a version number suffix, you create an object in the most recent version of PowerPoint that is available on the computer where the macro is running. For early binding, set a reference to  **Microsoft PowerPoint 15.0 Object Library**.



|**To create this object**|**Use one of these identifiers**|
|:-----|:-----|
|Application|PowerPoint.Application, PowerPoint.Application.15|
 **Microsoft Word**

To create the Microsoft Word objects listed in the following table, use one of the corresponding OLE programmatic identifiers. If you use an identifier without a version number suffix, you create an object in the most recent version of Word that is available on the computer where the macro is running. Word.Document.8 and Word.Document.12 both create a document in the default Open XML format (.docx). For early binding, set a reference to  **Microsoft Word 15.0 Object Library**.



|**To create this object**|**Use one of these identifiers**|
|:-----|:-----|
|Application|Word.Application, Word.Application.14|
|Document|Word.Document, Word.Document.8, Word.Template.8, Word.Document.12, Word.Template.12|

