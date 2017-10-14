---
title: Create or Replace a Worksheet
ms.prod: excel
ms.assetid: 227df739-3e66-4d23-8168-da43f552fbe0
ms.date: 06/08/2017
---


# Create or Replace a Worksheet

The following examples show how to determine if a worksheet exists, and then how to create or replace the worksheet.

 **Sample code provided by:** Tom Urtis, [Atlas Programming Management](http://www.atlaspm.com/)

## Determining if a Worksheet Exists

This example shows how to determine if a worksheet named "Sheet4" exists by using the  **[Name](worksheet-name-property-excel.md)** property of the **[Worksheet](worksheet-object-excel.md)** object. The name of the worksheet is specified by the `mySheetName` variable.


```vb
Sub TestSheetYesNo()
    Dim mySheetName As String, mySheetNameTest As String
    mySheetName = "Sheet4"
    
    On Error Resume Next
    mySheetNameTest = Worksheets(mySheetName).Name
    If Err.Number = 0 Then
        MsgBox "The sheet named ''" &; mySheetName &; "'' DOES exist in this workbook."
    Else
        Err.Clear
        MsgBox "The sheet named ''" &; mySheetName &; "'' does NOT exist in this workbook."
    End If
End Sub
```


## Creating the Worksheet

This example shows how to determine if a worksheet named "Sheet4" exists. The name of the worksheet is specified by the  `mySheetName` variable. If the worksheet does not exist, this example shows how to create a worksheet named "Sheet4" by using the **[Add](worksheets-add-method-excel.md)** method of the **[Worksheets](worksheets-object-excel.md)** object.


```vb
Sub TestSheetCreate()
    Dim mySheetName As String, mySheetNameTest As String
    mySheetName = "Sheet4"
    
    On Error Resume Next
    mySheetNameTest = Worksheets(mySheetName).Name
    If Err.Number = 0 Then
        MsgBox "The sheet named ''" &; mySheetName &; "'' DOES exist in this workbook."
    Else
        Err.Clear
        Worksheets.Add.Name = mySheetName
        MsgBox "The sheet named ''" &; mySheetName &; "'' did not exist in this workbook but it has been created now."
    End If
End Sub
```


## Replacing the Worksheet

This example shows how to determine if a worksheet named "Sheet4" exists. The name of the worksheet is specified by the  `mySheetName` variable. If the worksheet does exist, this example shows how to delete the existing worksheet by using the **[Delete](worksheet-delete-method-excel.md)** method of the **Worksheet** object, and then creates a new worksheet named "Sheet4".


 **Important**  All the data on the original worksheet named "Sheet4" is deleted when the worksheet is deleted.


```vb
Sub TestSheetReplace()
    Dim mySheetName As String
    mySheetName = "Sheet4"
    
    Application.DisplayAlerts = False
    On Error Resume Next
    Worksheets(mySheetName).Delete
    Err.Clear
    Application.DisplayAlerts = True
    Worksheets.Add.Name = mySheetName
    MsgBox "The sheet named ''" &; mySheetName &; "'' has been replaced."
End Sub
```


## About the Contributor
<a name="AboutContributor"> </a>

MVP Tom Urtis is the founder of Atlas Programming Management, a full-service Microsoft Office and Excel business solutions company in Silicon Valley. Tom has over 25 years of experience in business management and developing Microsoft Office applications, and is the coauthor of "Holy Macro! It's 2,500 Excel VBA Examples." 


