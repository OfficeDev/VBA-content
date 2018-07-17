---
title: Refer to Named Ranges
keywords: vbaxl10.chm5204437
f1_keywords:
- vbaxl10.chm5204437
ms.prod: excel
ms.assetid: 74119715-2208-b932-f47c-7fad334c3fc6
ms.date: 06/08/2017
---


# Refer to Named Ranges

Ranges are easier to identify by name than by A1 notation. To name a selected range, click the name box at the left end of the formula bar, type a name, and then press ENTER.

**Note**  There are two types of named ranges: Workbook Named Range and WorkSHEET Specific Named Range.   

## Workbook Named Range

A Workbook Named Range references a specific range from anywhere in the workbook (it applies globally).  

**How to Create a Workbook Named Range:**

As explained above, it is usually created entering the name into the name box to the left end of the formula bar. Note that no spaces are allowed in the name.  

## WorkSHEET Specific Named Range   

A WorkSHEET Specific Named Range refers to a range in a specific worksheet, and it is not global to all worksheets within a workbook. You can refer to this named range by just the name in the same worksheet, but from another worksheet you must use the worksheet name including "!"  the name of the range (example: the range "Name" "=Sheet1!Name"). 

The benefit is that you can use VBA code to generate new sheets with the same names for the same ranges within those sheets without getting an error saying that the name is already taken.   

**How to Create a WorkSHEET Specific Named Range:** 

1. Select the range you want to name.  
2. Click on the "Formulas" tab on the Excel Ribbon at the top of the window.   
3. Click "Define Name" button in the Formula tab.  
4. In the "New Name" dialogue box, under the field "Scope" choose the specific worksheet that the range you want to define is located (i.e. "Sheet1")- This makes the name specific to this worksheet. If you choose "Workbook" then it will be a WorkBOOK name).  

Example, of WorkSHEET Specific Named Range:  Selected range to name are A1:A10  

Chosen name of range is "name" within the same worksheet refer to the named name mere by entering the following in a cell "=name", from a different worksheet refer to the worksheet specific range by included the worksheet name in a cell "=Sheet1!name".  

## Referring to a Named Range

The following example refers to the range named "MyRange" in the workbook named "MyBook.xls."

```vba
Sub FormatRange() 
    Range("MyBook.xls!MyRange").Font.Italic = True 
End Sub
```

The following example refers to the worksheet-specific range named "Sheet1!Sales" in the workbook named "Report.xls."

```vba
Sub FormatSales() 
    Range("[Report.xls]Sheet1!Sales").BorderAround Weight:=xlthin 
End Sub
```

To select a named range, use the  **GoTo** method, which activates the workbook and the worksheet and then selects the range.

```vba
Sub ClearRange() 
    Application.Goto Reference:="MyBook.xls!MyRange" 
    Selection.ClearContents 
End Sub
```

The following example shows how the same procedure would be written for the active workbook.

```vba
Sub ClearRange() 
    Application.Goto Reference:="MyRange" 
    Selection.ClearContents 
End Sub
```

 **Sample code provided by:** Dennis Wallentin, [VSTO &; .NET &; Excel](http://xldennis.wordpress.com/)

This example uses a named range as the formula for data validation. This example requires the validation data to be on Sheet 2 in the range A2:A100. This validation data is used to validate data entered on Sheet 1 in the range D2:D10.

```vba
Sub Add_Data_Validation_From_Other_Worksheet()
'The current Excel workbook and worksheet, a range to define the data to be validated, and the target range
'to place the data in.
Dim wbBook As Workbook
Dim wsTarget As Worksheet
Dim wsSource As Worksheet
Dim rnTarget As Range
Dim rnSource As Range

'Initialize the Excel objects and delete any artifacts from the last time the macro was run.
Set wbBook = ThisWorkbook
With wbBook
    Set wsSource = .Worksheets("Sheet2")
    Set wsTarget = .Worksheets("Sheet1")
    On Error Resume Next
    .Names("Source").Delete
    On Error GoTo 0
End With

'On the source worksheet, create a range in column A of up to 98 cells long, and name it "Source".
With wsSource
    .Range(.Range("A2"), .Range("A100").End(xlUp)).Name = "Source"
End With

'On the target worksheet, create a range 8 cells long in column D.
Set rnTarget = wsTarget.Range("D2:D10")

'Clear out any artifacts from previous macro runs, then set up the target range with the validation data.
With rnTarget
    .ClearContents
    With .Validation
        .Delete
        .Add Type:=xlValidateList, _
             AlertStyle:=xlValidAlertStop, _
             Formula1:="=Source"
        
'Set up the Error dialog with the appropriate title and message
        .ErrorTitle = "Value Error"
        .ErrorMessage = "You can only choose from the list."
    End With
End With

End Sub
```

## Looping Through Cells in a Named Range

The following example loops through each cell in a named range by using a  **For Each...Next** loop. If the value of any cell in the range exceeds the value of `Limit`, the cell color is changed to yellow.

```vba
Sub ApplyColor() 
    Const Limit As Integer = 25 
    For Each c In Range("MyRange") 
        If c.Value > Limit Then 
            c.Interior.ColorIndex = 27 
        End If 
    Next c 
End Sub
```

## About the Contributor
<a name="AboutContributor"> </a>

Dennis Wallentin is the author of VSTO &; .NET &; Excel, a blog that focuses on .NET Framework solutions for Excel and Excel Services. Dennis has been developing Excel solutions for over 20 years and is also the coauthor of "Professional Excel Development: The Definitive Guide to Developing Applications Using Microsoft Excel, VBA and .NET (2nd Edition)." 


