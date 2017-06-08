---
title: Project.CurrencySymbol Property (Project)
keywords: vbapj.chm131697
f1_keywords:
- vbapj.chm131697
ms.prod: project-server
api_name:
- Project.Project.CurrencySymbol
ms.assetid: 5eccebc5-5c3d-4b30-31e0-68036411bca7
ms.date: 06/08/2017
---


# Project.CurrencySymbol Property (Project)

Gets or sets the characters that denote currency values. Read/write  **String**.


## Syntax

 _expression_. **CurrencySymbol**

 _expression_ A variable that represents a **Project** object.


## Remarks

Project sets the  **CurrencySymbol** property equal to the corresponding value in the **Customize Regional Options** dialog box of the Windows Control Panel.


## Example

The following example formats currency values in the active project according to the country or region specified by the user.


```vb
Sub FormatCurrency() 
 
    Dim CountryOrRegion As String 
 
    ' Prompt the user to enter the name of a country or region. 
    CountryOrRegion = UCase(InputBox$("Enter the name of a country or region: ", "Format Currency By Country Or Region")) 
     
    Select Case CountryOrRegion 
        Case "US", "United States", "USA", "United States of America" 
            ActiveProject.CurrencySymbol = "$" 
            ActiveProject.CurrencySymbolPosition = pjBefore 
        Case "ENGLAND" 
            ActiveProject.CurrencySymbol = Chr(163) 
            ActiveProject.CurrencySymbolPosition = pjBefore 
        Case "SWEDEN" 
            ActiveProject.CurrencySymbol = "kr" 
            ActiveProject.CurrencySymbolPosition = pjAfterWithSpace 
        ' Warn user if the currency format is not known. 
        Case Else 
            MsgBox ("The currency format for that country or region is unknown.") 
    End Select
End Sub
```


