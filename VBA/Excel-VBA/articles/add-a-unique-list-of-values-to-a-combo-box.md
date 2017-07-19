---
title: Add a Unique List of Values to a Combo Box
ms.prod: excel
ms.assetid: e2fa08b1-99bd-49fa-b1a2-5b693f7015e7
ms.date: 06/08/2017
---


# Add a Unique List of Values to a Combo Box

These examples show different approaches for taking a list from a spreadsheet and using it to populate a combo box control using only the unique values. The first example uses the  **AdvancedFilter** method of the Range object and the second uses the Collection object.

 **Sample code provided by:** Dennis Wallentin, [VSTO &; .NET &; Excel](http://xldennis.wordpress.com/)



```vb
Sub Populate_Combobox_Worksheet()
    'The Excel workbook and worksheets that contain the data, as well as the range placed on that data
    Dim wbBook As Workbook
    Dim wsSheet As Worksheet
    Dim rnData As Range

    'Variant to contain the data to be placed in the combo box.
    Dim vaData As Variant

    'Initialize the Excel objects
    Set wbBook = ThisWorkbook
    Set wsSheet = wbBook.Worksheets("Sheet1")

    'Set the range equal to the data, and then (temporarily) copy the unique values of that data to the L column.
    With wsSheet
        Set rnData = .Range(.Range("A1"), .Range("A100").End(xlUp))
        rnData.AdvancedFilter Action:=xlFilterCopy, _
                          CopyToRange:=.Range("L1"), _
                          Unique:=True
        'store the unique values in vaData
        vaData = .Range(.Range("L2"), .Range("L100").End(xlUp)).Value
        'clean up the contents of the temporary data storage
        .Range(.Range("L1"), .Range("L100").End(xlUp)).ClearContents
    End With

    'display the unique values in vaData in the combo box already in existence on the worksheet.
    With wsSheet.OLEObjects("ComboBox1").Object
        .Clear
        .List = vaData
        .ListIndex = -1
    End With

End Sub
```




```vb
Sub Populate_Combobox_Worksheet_Collection()
    'The Excel workbook and worksheets that contain the data, as well as the range placed on that data
    Dim wbBook As Workbook
    Dim wsSheet As Worksheet
    Dim rnData As Range

    Dim vaData As Variant               'the list, stored in a variant

    Dim ncData As New VBA.Collection    'the list, stored in a collection

    Dim lnCount As Long                 'the count used in the On Error Resume Next loop.

    Dim vaItem As Variant               'a variant representing the type of items in ncData

    'Instantiate the Excel objects.
    Set wbBook = ThisWorkbook
    Set wsSheet = wbBook.Worksheets("Sheet2")

    'Using Sheet2,retrieve the range of the list in Column A.
    With wsSheet
        Set rnData = .Range(.Range("A2"), .Range("A100").End(xlUp))
    End With

    'Place the list values into vaData.
    vaData = rnData.Value

    'Place the list values from vaData into the VBA.Collection.
    On Error Resume Next
        For lnCount = 1 To UBound(vaData)
        ncData.Add vaData(lnCount, 1), CStr(vaData(lnCount, 1))
    Next lnCount
    On Error GoTo 0
    
    'Clear the combo box (in case you ran the macro before),
    'and then add each unique variant item from ncData to the combo box.
    With wsSheet.OLEObjects("ComboBox1").Object
        .Clear
        For Each vaItem In ncData
            .AddItem ncData(vaItem)
        Next vaItem
    End With

End Sub
```


## About the Contributor
<a name="AboutContributor"> </a>

Dennis Wallentin is the author of VSTO &; .NET &; Excel, a blog that focuses on .NET Framework solutions for Excel and Excel Services. Dennis has been developing Excel solutions for over 20 years and is also the coauthor of "Professional Excel Development: The Definitive Guide to Developing Applications Using Microsoft Excel, VBA and .NET (2nd Edition)." 


