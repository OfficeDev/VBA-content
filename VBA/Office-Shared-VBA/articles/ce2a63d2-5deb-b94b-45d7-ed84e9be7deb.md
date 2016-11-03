
# ServerPolicy Object (Office)

Represents a policy specified for a document type stored on a server running Microsoft Office SharePoint Server.


## Remarks

The  **ServerPolicy** object is composed of individual **PolicyItem** objects representing the individual policy definitions for the active document.


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
|[Application](0d07cae1-9219-c617-f15d-01bc5ec59132.md)|
|[BlockPreview](a211ccbe-ee3e-168f-1f2f-15a1eddc876d.md)|
|[Count](aeb054d5-0b24-37e8-e1b6-6762a0d13d28.md)|
|[Creator](4acaac16-3611-ae19-9c6c-347ee67f6488.md)|
|[Description](ca820f97-79f7-d9aa-5368-e4ecfbfeccd3.md)|
|[Id](b1838ff9-d01a-bf19-a9a1-66627242eacc.md)|
|[Item](21fcec13-238e-f24d-2582-4c2ed8341d82.md)|
|[Name](a2afd663-55a0-913d-dade-19df4a1ab8dd.md)|
|[Parent](cab80a1e-f5e0-232f-c75b-14277f8a9022.md)|
|[Statement](7ae6f51a-bd5b-0a27-4a38-b07ff5c0d233.md)|

## See also


#### Other resources


[Object Model Reference](http://msdn.microsoft.com/library/499c789a-aba2-0fad-649a-0ea964cd3b5e%28Office.15%29.aspx)