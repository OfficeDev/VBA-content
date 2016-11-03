
# PolicyItem Object (Office)

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


## Properties



|**Name**|
|:-----|
|[Application](08c7efa1-7675-a327-67e1-db4f78fdf286.md)|
|[Creator](cef768a9-8c16-25dd-a596-7a9d2aa85bc3.md)|
|[Data](4ffa8c3a-f5fc-1813-daed-ea93f11df2dc.md)|
|[Description](3eaa6a5a-0606-5f1d-9ead-f7d92328173f.md)|
|[Id](b94f1822-78c9-ecad-e11b-002eae5e9762.md)|
|[Name](73dd5470-a229-d4a3-ded1-9821693e1a2a.md)|
|[Parent](280c24b7-bcab-4f61-ad10-e7cf13d47dd5.md)|

## See also


#### Other resources


[Object Model Reference](http://msdn.microsoft.com/library/499c789a-aba2-0fad-649a-0ea964cd3b5e%28Office.15%29.aspx)