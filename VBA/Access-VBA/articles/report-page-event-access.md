---
title: Report.Page Event (Access)
keywords: vbaac10.chm13882
f1_keywords:
- vbaac10.chm13882
ms.prod: access
api_name:
- Access.Report.Page
ms.assetid: c3fcce28-0bcd-4ef1-427f-504f0f80d336
ms.date: 06/08/2017
---


# Report.Page Event (Access)

The  **Page** event occurs after Microsoft Access formats a page of a report for printing, but before the page is printed. You can use this event to draw a border around the page, or add other graphic elements to the page.


## Syntax

 _expression_. **Page**

 _expression_ A variable that represents a **Report** object.


### Return Value

Nothing


## Remarks

To run a macro or event procedure when this event occurs, set the  **OnPage** property to the name of the macro or to [Event Procedure].

This event occurs after all the  **Format** events for the report, and after all the **Print** events for the page, but before the page is actually printed.

You normally use the  **Line**, **Circle**, or **PSet** methods in the **Page** event procedure to create the desired graphics for the page.

The  **NoData** event occurs before the first **Page** event for the report.


## Example

The following example shows how to draw a rectangle around a report page by using the  **Line** method. The **ScaleWidth** and **ScaleHeight** properties by default return the internal width and height of the report.


```vb
Private Sub Report_Page() 
    Me.Line(0, 0)-(Me.ScaleWidth, Me.ScaleHeight), , B 
End Sub
```

The following example shows how to use the  **Page** event to add a watermark to a report before it is printed.


 **Sample code provided by:** The[Microsoft Access 2010 Programmer's Reference](http://www.wrox.com/WileyCDA/WroxTitle/Access-2010-Programmer-s-Reference.productCd-0470591668.html)





```vb
Private Sub Report_Page()
    Dim strWatermarkText As String
    Dim sizeHor As Single
    Dim sizeVer As Single

#If RUN_PAGE_EVENT = True Then
    With Me
        '// Print page border
        Me.Line (0, 0)-(.ScaleWidth - 1, .ScaleHeight - 1), vbBlack, B
    
        '// Print watermark
        strWatermarkText = "Confidential"
        
        .ScaleMode = 3
        .FontName = "Segoe UI"
        .FontSize = 48
        .ForeColor = RGB(255, 0, 0)

        '// Calculate text metrics
        sizeHor = .TextWidth(strWatermarkText)
        sizeVer = .TextHeight(strWatermarkText)
        
        '// Set the print location
        .CurrentX = (.ScaleWidth / 2) - (sizeHor / 2)
        .CurrentY = (.ScaleHeight / 2) - (sizeVer / 2)
    
        '// Print the watermark
        .Print strWatermarkText
    End With
#End If

End Sub
```


## About the Contributors
<a name="AboutContributors"> </a>

Wrox Press is driven by the Programmer to Programmer philosophy. Wrox books are written by programmers for programmers, and the Wrox brand means authoritative solutions to real-world programming problems. 


## See also
<a name="AboutContributors"> </a>


#### Concepts


[Report Object](report-object-access.md)

