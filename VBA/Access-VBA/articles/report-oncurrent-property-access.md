---
title: Report.OnCurrent Property (Access)
keywords: vbaac10.chm13823
f1_keywords:
- vbaac10.chm13823
ms.prod: access
api_name:
- Access.Report.OnCurrent
ms.assetid: 593fdb6c-017a-986f-22ef-cc9e66aaaf01
ms.date: 06/08/2017
---


# Report.OnCurrent Property (Access)

Sets or returns the value of the  **On Current** property on the Report. Read/write **String**.


## Syntax

 _expression_. **OnCurrent**

 _expression_ A variable that represents a **Report** object.


## Remarks

If you want a procedure to run automatically every time you open a particular report, you set the form's  **On Current** property to _"[Event Procedure]"_ and Access automatically stubs-out a procedure for you called _Private Sub Report_Current()_. The **OnCurrent** property allows you to programmatically determine the value of the form's **On Current** property, or to programmatically set the form's **On Current** property.

 _Note that the  **Current** event fires when you run (open) a report._

If you set the form's  **On Current** property in the UI, it gets it value based on your selection in the **Choose Builder** window, which appears when you click the **...** button next to **On Current** in the Report's **Properties** window.


- If you select  **Code Builder**, then the value will be  _[Event Procedure]_.
    
- If you select  **Expression Builder**, then the value will be "= _expression_ ", where _expression_ is the expression from the Expression Builder window.
    
- If you select  **Macro Builder**, then the value will be the name of the macro.
    

## Example

The following code example demonstrates how to set a Report's  **OnCurrent** property.


```vb

Private Sub Report_Load()

        Me.OnCurrent = "[Event Procedure]"

End Sub
		
```

The event procedure  **Report_Current()** is automatically called when the **Current** event is fired. This procedure simply collects the values of two of the Report's text boxes and sends them to another procedure for processing.




```vb

Private Sub Report_Current()

        ' Declare variables to store price and available credit.
        Dim curPrice As Currency
        Dim curCreditAvail As Currency

        ' Assign variables from current values in text boxes on the Report.
        curPrice = txtValue1
        curCreditAvail = txtValue2

        ' Call VerifyCreditAvail procedure.
        VerifyCreditAvail curPrice, curCreditAvail

End Sub
		
```

The following code example simply processes the two values passed to it.




```vb
Sub VerifyCreditAvail(curTotalPrice As Currency, curAvailCredit As Currency)
    ' Inform the user if there is not enough credit available for the purchase.
    If curTotalPrice > curAvailCredit Then
        MsgBox "You do not have enough credit available for this purchase."
    End If
End Sub
```


## See also


#### Concepts


[Report.Current Event (Access)](report-current-event-access.md)
[Report Object](report-object-access.md)

