---
title: Hide and Unhide Columns
ms.prod: excel
ms.assetid: fbfd24bb-9862-4895-9ac4-3e4f92197ede
ms.date: 06/08/2017
---


# Hide and Unhide Columns

This example finds all the cells in the first four columns that have a constant &;quot;X&;quot; in them and hides the column that contains the X.

 **Sample code provided by:** Dennis Wallentin,[VSTO &; .NET &; Excel](http://xldennis.wordpress.com/)



```
Sub Hide_Columns()

    'Excel objects.
    Dim m_wbBook As Workbook
    Dim m_wsSheet As Worksheet
    Dim m_rnCheck As Range
    Dim m_rnFind As Range
    Dim m_stAddress As String

    'Initialize the Excel objects.
    Set m_wbBook = ThisWorkbook
    Set m_wsSheet = m_wbBook.Worksheets("Sheet1")
    
    'Search the four columns for any constants.
    Set m_rnCheck = m_wsSheet.Range("A1:D1").SpecialCells(xlCellTypeConstants)
    
    'Retrieve all columns that contain an X. If there is at least one, begin the DO/WHILE loop.
    With m_rnCheck
        Set m_rnFind = .Find(What:="X")
        If Not m_rnFind Is Nothing Then
            m_stAddress = m_rnFind.Address
             
            'Hide the column, and then find the next X.
            Do
                m_rnFind.EntireColumn.Hidden = True
                Set m_rnFind = .FindNext(m_rnFind)
            Loop While Not m_rnFind Is Nothing And m_rnFind.Address <> m_stAddress
        End If
    End With

End Sub
```

This example finds all the cells in the first four columns that have a constant "X" in them and unhides the column that contains the X.



```
Sub Unhide_Columns()
    'Excel objects.
    Dim m_wbBook As Workbook
    Dim m_wsSheet As Worksheet
    Dim m_rnCheck As Range
    Dim m_rnFind As Range
    Dim m_stAddress As String
    
    'Initialize the Excel objects.
    Set m_wbBook = ThisWorkbook
    Set m_wsSheet = m_wbBook.Worksheets("Sheet1")
    
    'Search the four columns for any constants.
    Set m_rnCheck = m_wsSheet.Range("A1:D1").SpecialCells(xlCellTypeConstants)
    
    'Retrieve all columns that contain X. If there is at least one, begin the DO/WHILE loop.
    With m_rnCheck
        Set m_rnFind = .Find(What:="X", LookIn:=xlFormulas)
        If Not m_rnFind Is Nothing Then
            m_stAddress = m_rnFind.Address
            
            'Unhide the column, and then find the next X.
            Do
                m_rnFind.EntireColumn.Hidden = False
                Set m_rnFind = .FindNext(m_rnFind)
            Loop While Not m_rnFind Is Nothing And m_rnFind.Address <> m_stAddress
        End If
    End With

End Sub
```


## About the Contributor
<a name="AboutContributor"> </a>

Dennis Wallentin is the author of VSTO &; .NET &; Excel, a blog that focuses on .NET Framework solutions for Excel and Excel Services. Dennis has been developing Excel solutions for over 20 years and is also the coauthor of "Professional Excel Development: The Definitive Guide to Developing Applications Using Microsoft Excel, VBA and .NET (2nd Edition)." 


