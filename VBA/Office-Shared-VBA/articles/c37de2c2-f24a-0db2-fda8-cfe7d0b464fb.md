
# SmartDocument.RefreshPane Method (Office)

Refreshes the  **Document Actions** task pane for the active document in Microsoft Word or a workbook in Microsoft Excel.


## Syntax

 _expression_. **RefreshPane**

 _expression_ A variable that represents a **SmartDocument** object.


## Remarks

The  **RefreshPane** method raises an error if the active document does not have an XML expansion pack attached.


## Example

The following example determines whether the active Excel workbook has an XML expansion pack attached. If so, it refreshes the smart document's  **Document Actions** task pane.


```
 Dim objSmartDoc As Office.SmartDocument 
 Set objSmartDoc = ActiveWorkbook.SmartDocument 
 If objSmartDoc.SolutionID > "None" Then 
 objSmartDoc.RefreshPane 
 Else 
 MsgBox "No XML expansion pack attached." 
 End If 

```


## See also


#### Concepts


[SmartDocument Object](b56a86eb-a031-d50b-905e-ef8b91914d61.md)
#### Other resources


[SmartDocument Object Members](980de42d-6992-6107-a3fb-33e8c78da202.md)
