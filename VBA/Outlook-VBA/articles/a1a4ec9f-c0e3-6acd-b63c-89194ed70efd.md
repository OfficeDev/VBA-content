
# Application.AdvancedSearchStopped Event (Outlook)

Occurs when a specified  **[Search](226a5d49-3caf-90dd-725c-265404d1939f.md)** object's **[Stop](c087e5aa-a846-56e1-a808-e8718096c3c9.md)** method has been executed.


## Syntax

 _expression_ . **AdvancedSearchStopped**( **_SearchObject_** )

 _expression_ A variable that represents an **Application** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _SearchObject_|Required| **Search**|The  **[Search](226a5d49-3caf-90dd-725c-265404d1939f.md)** object returned by the **[AdvancedSearch](7b433d8b-08b9-dff1-b854-287d76b47a90.md)** method.|

## Remarks

After this event is fired, the  **Search** object?s **[Results](59057f6f-8f6d-eed0-c945-240b9593b7ea.md)** collection will no longer be updated. This event can only be triggered programmatically.


## Example

The following Visual Basic for Applications (VBA) example starts searching the  **Inbox** for items with subject equal to "Test" and immediately stops the search. This causes the `AdvanceSearchStopped` event procedure to be run. The sample code must be placed in a class module such as `ThisOutlookSession`. The  `StopSearch()` procedure must be called before the event procedure can be called by Microsoft Outlook.


```vb
Sub StopSearch() 
 
 Dim sch As Outlook.Search 
 
 Dim strScope As String 
 
 Dim strFilter As String 
 
 strScope = "Inbox" 
 
 strFilter = "urn:schemas:httpmail:subject = 'Test'" 
 
 Set sch = Application.AdvancedSearch(strScope, strFilter) 
 
 sch.Stop 
 
End Sub 
 
 
 
Private Sub Application_AdvancedSearchStopped(ByVal SearchObject As Search) 
 
 'Inform the user that the search has stopped. 
 
 MsgBox "An AdvancedSearch has been interrupted and stopped. " 
 
End Sub
```


## See also


#### Concepts


[Application Object](797003e7-ecd1-eccb-eaaf-32d6ddde8348.md)
#### Other resources


[Application Object Members](3519c89c-2353-85ee-7ddc-62e5dd85a8e7.md)
