---
title: Find All the Sparklines on a Sheet
ms.prod: excel
ms.assetid: 39739eaf-638d-41b1-80f2-c4513fc42317
ms.date: 06/08/2017
---


# Find All the Sparklines on a Sheet

The following code example uses a list box on a user form to display all of the sparkline groups on the active sheet. When you click one of the sparkline groups in the list box, the sparkline group is selected on the sheet.

This example requires a user form named  **SparklineForm**, a list box on the user form named  **SparklineListBox**, and a button on the user form named  **CloseBtn**.

In the Visual Basic Editor, insert a  **Module** and copy and paste the following code. This code shows the user form.




```vb
Sub ShowUserForm()
    SparklineForm.Show
End Sub
```

In the Visual Basic Editor, right-click the  **SparklineForm** form, select **View Code**, and copy and paste the following code.
The  **UserForm_Activate** procedure iterates through all the sparkline groups on the active sheet and gets the addresses of the sparkline groups by using the [Address](range-address-property-excel.md) property of the [Range](range-object-excel.md) object. The address is then added to the list box.
The  **SparklineListBox_Click** procedure is called when you click the address of a sparkline group in the list box. This procedure activates the selected sparkline group on the sheet by using the [Activate](range-activate-method-excel.md) method of the [Range](range-object-excel.md) object.
The  **CloseBtn_Click** procedure is called when you click the button on the user form, and it closes the user form.



```vb
Private Sub UserForm_Activate()
    'The sparkline group
    Dim oSparkGroup As SparklineGroup
    
    'Loop through all the sparkline groups on the sheet
    For Each oSparkGroup In ActiveSheet.Range("A:XFD").SparklineGroups
        'For each sparkline group found, add the address to the listbox
        SparklineListBox.AddItem oSparkGroup.Location.Address(, , , True)
    Next oSparkGroup
End Sub

Private Sub SparklineListBox_Click()
    'Activate the selected range that has the sparklines
    Range(SparklineListBox.Value).Activate
End Sub

Private Sub CloseBtn_Click()
    'Close the userform
    Unload Me
End Sub
```


## See also


#### Concepts


 [SparklineGroup Object](sparklinegroup-object-excel.md)
#### Other resources


 <br>
 [Programming With Sparklines In Excel](http://msdn.microsoft.com/library/e26f3356-882e-44d5-94a5-c7e8d1026d78%28Office.15%29.aspx)

