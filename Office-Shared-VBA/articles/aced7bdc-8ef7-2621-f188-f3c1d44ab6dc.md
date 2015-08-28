
# PolicyItem Object (Office)

 **Last modified:** July 28, 2015

Represents an item within a  **ServerPolicy** object that contains the settings for one policy.

## Remarks

A policy item cannot exist outside the scope of a policy. Policy items are distinct conditions defined for a document stored on a server running Microsoft Office SharePoint Server.


## Example

The following example lists the name and description of all of the policy items for the active document.


```
Sub ListPolicyItems() 
Dim objSrvPolicy As ServerPolicy 
Dim objPolicyItem As PolicyItem 
Dim strPolicyItemList As String 
 
Set objSrvPolicy = ActiveDocument.ServerPolicy 
 
For Each objPolicyItem In objSrvPolicy 
 strPolicyItemList = "Policy Item " &amp; objPolicyItem.Name &amp; " - " &amp; _ 
 objPolicyItem.Description &amp; vbCrLf 
Next 
MsgBox (strPolicyItemList) 
 
End Sub 

```


## See also


#### Concepts


 [Object Model Reference](499c789a-aba2-0fad-649a-0ea964cd3b5e.md)
#### Other resources


 [PolicyItem Object Members](a2e43e08-64bb-f052-78a2-0618e2df46fc.md)
