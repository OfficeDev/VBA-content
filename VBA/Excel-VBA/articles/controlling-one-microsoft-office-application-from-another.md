---
title: Controlling One Microsoft Office Application from Another
keywords: vbaxl10.chm5200116
f1_keywords:
- vbaxl10.chm5200116
ms.prod: excel
ms.assetid: 588c18f7-b9e4-60df-e209-a411c5a22fc6
ms.date: 06/08/2017
---


# Controlling One Microsoft Office Application from Another

If you want to run code in one Microsoft Office application that works with the objects in another application, follow these steps.


### To run the code


1. Set a reference to the other application's type library in the  **References** dialog box ( **Tools** menu). Then, the objects, properties, and methods will appear in the Object Browser and the syntax will be checked at compile time. You can also get context-sensitive Help on them.
    
2. Declare object variables that will refer to the objects in the other application as specific types. Qualify each type with the name of the application that is supplying the object. For example, the following statement declares a variable that points to a Microsoft Word document and another that refers to a Microsoft Excel workbook.
    
    ```vb
      Dim appWD As Word.Application, wbXL As Excel.Workbook
    ```
    
     **Note**  You must follow the preceding steps if you want your code to be early bound.
     
3. Use the  **CreateObject** function with the [OLE Programmatic Identifiers](http://msdn.microsoft.com/library/9d3418b1-cf9e-4c4d-c387-07952f41608e%28Office.15%29.aspx) of the object you want to work with in the other application, as shown in the following example. To see the session of the other application, set the **Visible** property to **True**.
        
    ```vb
      Dim appWD As Word.Application 
     
    Set appWD = CreateObject("Word.Application") 
    appWd.Visible = True
    ```
    
4. Apply properties and methods to the object contained in the variable. For example, the following instruction creates a new Word document.
        
    ```vb
    Dim appWD As Word.Application 
     
    Set appWD = CreateObject("Word.Application") 
    appWD.Documents.Add
    ```

5. When finished working with the other application, use the  **Quit** method to close it, and then set its object variable to **Nothing** to free any memory it is using, as shown in the following example.
    
    ```vb
    appWd.Quit 
    Set appWd = Nothing
    ```
    
 **Sample code provided by:** Bill Jelen, [MrExcel.com](http://www.mrexcel.com/)
The following code example creates a new Microsoft Office Word file for each row of data in a spreadsheet.
    
```vb
' You must pick Microsoft Word Object Library from Tools>References
' in the VB editor to execute Word commands.
Sub ControlWord()
    Dim appWD As Word.Application
    ' Create a new instance of Word and make it visible
    Set appWD = CreateObject("Word.Application.12")
    appWD.Visible = True

    ' Find the last row with data in the spreadsheet
    FinalRow = Range("A9999").End(xlUp).Row
    For i = 1 To FinalRow
        ' Copy the current row
        Worksheets("Sheet1").Rows(i).Copy
        ' Tell Word to create a new document
        appWD.Documents.Add
        ' Tell Word to paste the contents of the clipboard into the new document.
        appWD.Selection.Paste
        ' Save the new document with a sequential file name.
        appWD.ActiveDocument.SaveAs Filename:="File" &; i
        ' Close the new Word document.
        appWD.ActiveDocument.Close
    Next i
    ' Close the Word application.
    appWD.Quit
End Sub
```

**Sample code provided by:** Dennis Wallentin, [VSTO &; .NET &; Excel](http://xldennis.wordpress.com/)
This example takes the cells values from a named range,  **W_Data**, that contains three values and inserts those values into a Word document. The values are inserted at bookmarked locations named  **td1**,  **td2**, and  **td3**.
For this example to run, you must have a range named  **W_Data** that contains three values on **Sheet1** in the workbook. You must have a Word document named **Test.docx** saved in the same location as the Excel workbook, and the Word document must have three bookmarks named **td1**,  **td2**, and  **td3**.

```vb
' You must pick Microsoft Word Object Library from Tools>References
' in the Visual Basic editor to execute Word commands.

Option Explicit

Sub Add_Single_Values_Word()
Dim wdApp As Word.Application
Dim wdDoc As Word.Document
Dim wdRange1 As Word.Range
Dim wdRange2 As Word.Range
Dim wdRange3 As Word.Range

Dim wbBook As Workbook
Dim wsSheet As Worksheet

Dim vaData As Variant

Set wbBook = ThisWorkbook
Set wsSheet = wbBook.Worksheets("Sheet1")

vaData = wsSheet.Range("W_Data").Value

' Instatiate the Word Objects.
Set wdApp = New Word.Application
Set wdDoc = wdApp.Documents.Open(wbBook.Path &; "\Test.docx")

With wdDoc
    Set wdRange1 = .Bookmarks("td1").Range
    Set wdRange2 = .Bookmarks("td2").Range
    Set wdRange3 = .Bookmarks("td3").Range
End With

' Write values to the bookmarks.
wdRange1.Text = vaData(1, 1)
wdRange2.Text = vaData(2, 1)
wdRange3.Text = vaData(3, 1)

With wdDoc
    .Save
    .Close
End With

wdApp.Quit

' Release the objects from memory.
Set wdRange1 = Nothing
Set wdRange2 = Nothing
Set wdRange3 = Nothing
Set wdDoc = Nothing
Set wdApp = Nothing

End Sub
```

## About the Contributors

<a name="AboutContributor"> </a>

MVP Bill Jelen is the author of more than two dozen books about Microsoft Excel. He is a regular guest on TechTV with Leo Laporte and is the host of MrExcel.com, which includes more than 300,000 questions and answers about Excel. 

Dennis Wallentin is the author of VSTO &; .NET &; Excel, a blog that focuses on .NET Framework solutions for Excel and Excel Services. Dennis has been developing Excel solutions for over 20 years and is also the coauthor of "Professional Excel Development: The Definitive Guide to Developing Applications Using Microsoft Excel, VBA and .NET (2nd Edition)." 

