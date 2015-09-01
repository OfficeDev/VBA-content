
# ColumnFormat Object (Outlook)

 **Last modified:** July 28, 2015

Represents the display properties of an order field or view field in a view.

## Remarks

The  **ColumnFormat** object represents the display properties, such as the alignment or field type, of an ** [OrderField](4ae32270-bde9-3178-bca3-f8d145779d3d.md)** or ** [ViewField](997319f0-7ff3-a712-8484-2e442965e187.md)** object. Use the ** [ColumnFormat](0014f1d8-5380-3301-558a-7fd8d49afff9.md)** property of the **ViewField** object to access the display properties of a view field.

Use the  ** [Label](cf104506-3eca-6695-3d3b-05022ce6fba4.md)** property to obtain or change the text used to label the field, or the ** [Align](cea9e062-e338-ee1d-f769-dd5f8beef463.md)** property to determine the alignment of the contents within the field.

Use the  ** [FieldType](84a40f6f-72fe-61e5-d85c-7a7c90f3e58a.md)** property to determine the type and form of the data displayed for that field, and the ** [FieldFormat](14064b56-65c2-1c7d-1e74-3bfa2d2ccaa7.md)** property to determine how to format the data for that field.


## Example

The following Visual Basic for Applications (VBA) example iterates through the  ** [ViewFields](c4c6257e-fdbe-c187-86c5-34bee3eb0bd3.md)** collection of the current ** [TableView](026e27f8-1655-060d-e8cc-87eaaf4f1510.md)** object, displaying the label and XML schema names of each **ViewField** object in the collection.


```
Private Sub DisplayTableViewFields() 
 
 Dim objTableView As TableView 
 
 Dim objViewField As ViewField 
 
 Dim strOutput As String 
 
 
 
 If Application.ActiveExplorer.CurrentView.ViewType = _ 
 
 olTableView Then 
 
 
 
 ' Obtain a TableView object reference for the 
 
 ' current table view. 
 
 Set objTableView = _ 
 
 Application.ActiveExplorer.CurrentView 
 
 
 
 ' Iterate through the ViewFields collection for 
 
 ' the table view, obtaining the label and the 
 
 ' XML schema name for each field included in 
 
 ' the view. 
 
 For Each objViewField In objTableView.ViewFields 
 
 With objViewField 
 
 strOutput = strOutput &amp; .ColumnFormat.Label &amp; _ 
 
 " (" &amp; .ViewXMLSchemaName &amp; ")" &amp; vbCrLf 
 
 End With 
 
 Next 
 
 
 
 ' Display a dialog box containing the concatenated 
 
 ' view field information. 
 
 MsgBox strOutput 
 
 End If 
 
End Sub 
 

```


## See also


#### Concepts


 [Outlook Object Model Reference](73221b13-d8d8-99b8-3394-b95dbbfd5ddc.md)
#### Other resources


 [ColumnFormat Object Members](7159f452-7a05-f3a3-53f8-0b3f5463d313.md)
