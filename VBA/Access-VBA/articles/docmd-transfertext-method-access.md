---
title: DoCmd.TransferText Method (Access)
keywords: vbaac10.chm4190
f1_keywords:
- vbaac10.chm4190
ms.prod: access
api_name:
- Access.DoCmd.TransferText
ms.assetid: e59f26dc-2df8-8d87-b73d-f3004eed0719
ms.date: 11/30/2017
---


# DoCmd.TransferText Method (Access)

The **TransferText** method carries out the **TransferText** action in Visual Basic.


## Syntax

_expression_. **TransferText** (**_TransferType_**, **_SpecificationName_**, **_TableName_**, **_FileName_**, **_HasFieldNames_**, **_HTMLTableName_**, **_CodePage_**)

_expression_ A variable that represents a **DoCmd** object.


### Parameters

|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
|_TransferType_|Optional|[AcTextTransferType](actexttransfertype-enumeration-access.md)|The type of transfer you want to make. You can import data from, export data to, or link to data in delimited or fixed-width text files or HTML files. The default value is  **acImportDelim**. Only **acImportDelim**, **acImportFixed**, **acExportDelim**, **acExportFixed**, or **acExportMerge** transfer types are supported in a Microsoft Access project (.adp).|
|_SpecificationName_|Optional|**Variant**|A string expression that's the name of an import or export specification you've created and saved in the current database. For a fixed-width text file, you must either specify an argument or use a schema.ini file, which must be stored in the same folder as the imported, linked, or exported text file. To create a schema file, you can use the text import/export wizard to create the file. For delimited text files and Microsoft Word mail merge data files, you can leave this argument blank to select the default import/export specifications.|
|_TableName_|Optional|**Variant**|A string expression that's the name of the Microsoft Access table you want to import text data to, export text data from, or link text data to, or the Microsoft Access query whose results you want to export to a text file.|
|_FileName_|Optional|**Variant**|A string expression that's the full name, including the path, of the text file you want to import from, export to, or link to.|
|_HasFieldNames_|Optional|**Variant**|Use  **True** (1) to use the first row of the text file as field names when importing, exporting, or linking. Use **False** (0) to treat the first row of the text file as normal data. If you leave this argument blank, the default ( **False** ) is assumed. This argument is ignored for Microsoft Word mail merge data files, which must always contain the field names in the first row.|
|_HTMLTableName_|Optional|**Variant**|A string expression that's the name of the table or list in the HTML file that you want to import or link. This argument is ignored unless the  _TransferType_ argument is set to **acImportHTML** or **acLinkHTML**. If you leave this argument blank, the first table or list in the HTML file is imported or linked. The name of the table or list in the HTML file is determined by the text specified by the **CAPTION** tag, if there's a **CAPTION** tag. If there's no **CAPTION** tag, the name is determined by the text specified by the **TITLE** tag. If more than one table or list has the same name, Microsoft Access distinguishes them by adding a number to the end of each table or list name; for example, Employees1 and Employees2.|
|[CodePage](https://msdn.microsoft.com/en-us/library/windows/desktop/dd317756(v=vs.85).aspx)|Optional|**Variant**|A **Long** value indicating the character set of the code page.|

## Remarks

You can use the  **TransferText** method to import or export text between the current Microsoft Access database or Access project (.adp) and a text file. You can also link the data in a text file to the current Access database. With a linked text file, you can view the text data with Access while still allowing complete access to the data from your word processing program. You can also import from, export to, and link to a table or list in an HTML file (*.html).

You can export the data in Access select queries to text files. Access exports the result set of the query, treating it just like a table.


## Example

The following example exports the data from the Microsoft Access table External Report to the delimited text file April.doc by using the specification Standard Output:

```vb
DoCmd.TransferText acExportDelim, "Standard Output", _ 
    "External Report", "C:\Txtfiles\April.doc"
```

The following code shows how to create a new Microsoft Word document and perform a mail merge with the data stored in the Customers table.

**Sample code provided by:** The [Microsoft Access 2010 Programmer's Reference](http://www.wrox.com/WileyCDA/WroxTitle/Access-2010-Programmer-s-Reference.productCd-0470591668.html)


```vb
Public Sub DoMailMerge(strFileSavePath As String)

    ' Create new Word App, add a document and set it visible
    Dim wdApp As New Word.Application
    wdApp.Documents.Add
    wdApp.Visible = True

    ' Open the data set from this database
    wdApp.ActiveDocument.MailMerge.OpenDataSource _
        Name:=Application.CurrentProject.FullName, _
        OpenExclusive:=False, _
        LinkToSource:=True, _
        Connection:="TABLE Customers", _
        SQLStatement:="SELECT Customers.* FROM Customers;"
              
    ' Add fields to the mail merge document
    Dim oSel As Object
    Set oSel = wdApp.Selection
    With wdApp.ActiveDocument.MailMerge.Fields
    
        oSel.TypeText vbNewLine &; vbNewLine
        .Add oSel.range, "First_Name"
        oSel.TypeText " "
        .Add oSel.range, "Last_Name"
        oSel.TypeText vbNewLine
        .Add oSel.range, "Company"
        oSel.TypeText vbNewLine
        .Add oSel.range, "Address"
        oSel.TypeText vbNewLine
        .Add oSel.range, "City"
        oSel.TypeText ", "
        .Add oSel.range, "State"
        oSel.TypeText " "
        .Add oSel.range, "Zip"
        oSel.TypeText vbNewLine
        oSel.TypeParagraph
        oSel.TypeText "Dear "
        .Add oSel.range, "First_Name"
        oSel.TypeText ","
        oSel.TypeText vbNewLine
        oSel.TypeParagraph
        oSel.TypeText "We have created this mail just for you..."
        oSel.TypeText vbNewLine
        oSel.TypeText vbNewLine
        oSel.TypeText "Sincerely," &; vbNewLine &; "John Q. Public"
        oSel.TypeText vbFormFeed
        
    End With
    
    ' Execute the mail merge and save the document
    wdApp.ActiveDocument.MailMerge.Execute
    wdApp.ActiveDocument.SaveAs strFileSavePath
        
    ' Close everything and Cleanup Variables
    Set oSel = Nothing
    wdApp.ActiveDocument.Close False
    Set wdApp = Nothing

End Sub
```


## About the contributors
<a name="AboutContributors"> </a>

Wrox Press is driven by the Programmer to Programmer philosophy. Wrox books are written by programmers for programmers, and the Wrox brand means authoritative solutions to real-world programming problems. 


## See also

[DoCmd Object](docmd-object-access.md)

