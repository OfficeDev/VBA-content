---
title: ErrorCheckingOptions Object (Excel)
keywords: vbaxl10.chm697072
f1_keywords:
- vbaxl10.chm697072
ms.prod: excel
api_name:
- Excel.ErrorCheckingOptions
ms.assetid: f62d3b08-a08f-d028-8e33-4bfd8799dc44
ms.date: 06/08/2017
---


# ErrorCheckingOptions Object (Excel)

Represents the error-checking options for an application.


## Remarks

Use the  **[ErrorCheckingOptions](application-errorcheckingoptions-property-excel.md)** property of the **[Application](application-object-excel.md)** object to return an **ErrorCheckingOptions** object.

Reference the  **[Item](errors-item-property-excel.md)** property of the **[Errors](errors-object-excel.md)** object to view a list of index values associated with error-checking options.

Once an  **ErrorCheckingOptions** object is returned, you can use the following properties, which are members of the **ErrorCheckingOptions** object, to set or return error checking options.


-  **[BackgroundChecking](errorcheckingoptions-backgroundchecking-property-excel.md)**
    
-  **[EmptyCellReferences](errorcheckingoptions-emptycellreferences-property-excel.md)**
    
-  **[EvaluateToError](errorcheckingoptions-evaluatetoerror-property-excel.md)**
    
-  **[InconsistentFormula](errorcheckingoptions-inconsistentformula-property-excel.md)**
    
-  **[IndicatorColorIndex](errorcheckingoptions-indicatorcolorindex-property-excel.md)**
    
-  **[NumberAsText](errorcheckingoptions-numberastext-property-excel.md)**
    
-  **[OmittedCells](errorcheckingoptions-omittedcells-property-excel.md)**
    
-  **[TextDate](errorcheckingoptions-textdate-property-excel.md)**
    
-  **[UnlockedFormulaCells](errorcheckingoptions-unlockedformulacells-property-excel.md)**
    

## Example

The following example uses the  **TextDate** property to enable error checking for two-digit-year text dates and notifies the user.


```vb
Sub CheckTextDates() 
 
 Dim rngFormula As Range 
 Set rngFormula = Application.Range("A1") 
 
 Range("A1").Formula = "'April 23, 00" 
 Application.ErrorCheckingOptions.TextDate = True 
 
 ' Perform check to see if 2 digit year TextDate check is on. 
 If rngFormula.Errors.Item(xlTextDate).Value = True Then 
 MsgBox "The text date error checking feature is enabled." 
 Else 
 MsgBox "The text date error checking feature is not on." 
 End If 
 
End Sub
```


## See also


#### Other resources



[Excel Object Model Reference](http://msdn.microsoft.com/library/11ea8598-8a20-92d5-f98b-0da04263bf2c%28Office.15%29.aspx)

