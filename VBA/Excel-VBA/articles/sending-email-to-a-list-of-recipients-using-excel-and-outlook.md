---
title: Sending Email to a List of Recipients Using Excel and Outlook
ms.prod: excel
ms.assetid: 207b0384-30f0-412a-8d2b-c1740fb61420
ms.date: 06/08/2017
---


# Sending Email to a List of Recipients Using Excel and Outlook

The following code example shows how to send an email to a list of recipients based on data stored in a workbook. The recipient email addresses must be in column A, and the body text of the email must be in the first text box on the active sheet.

 **Sample code provided by:** Holy Macro! Books, [Holy Macro! It's 2,500 Excel VBA Examples](http://www.mrexcel.com/store/index.php?l=product_detail&;p=1)



```vb
Sub Sample()
   'Setting up the Excel variables.
   Dim olApp As Object
   Dim olMailItm As Object
   Dim iCounter As Integer
   Dim Dest As Variant
   Dim SDest As String
   
   'Create the Outlook application and the empty email.
   Set olApp = CreateObject("Outlook.Application")
   Set olMailItm = olApp.CreateItem(0)
   
   'Using the email, add multiple recipients, using a list of addresses in column A.
   With olMailItm
       SDest = ""
       For iCounter = 1 To WorksheetFunction.CountA(Columns(1))
           If SDest = "" Then
               SDest = Cells(iCounter, 1).Value
           Else
               SDest = SDest &; ";" &; Cells(iCounter, 1).Value
           End If
       Next iCounter
       
    'Do additional formatting on the BCC and Subject lines, add the body text from the spreadsheet, and send.
       .BCC = SDest
       .Subject = "FYI"
       .Body = ActiveSheet.TextBoxes(1).Text
       .Send
   End With
   
   'Clean up the Outlook application.
   Set olMailItm = Nothing
   Set olApp = Nothing
End Sub
```


## About the Contributor
<a name="AboutContributor"> </a>

Holy Macro! Books publishes entertaining books for people who use Microsoft Office. See the complete catalog at MrExcel.com. 


