
# DocumentProperty.LinkToContent Property (Office)

 **Last modified:** July 28, 2015

 **In this article**
 [Syntax](#sectionSection0)
 [Remarks](#sectionSection1)
 [Example](#sectionSection2)


Is  **True** if the value of the custom document property is linked to the content of the container document. **False** if the value is static. Read/write.


## Syntax
<a name="sectionSection0"> </a>

 _expression_. **LinkToContent**( **_pfLinkRetVal_**)

 _expression_A variable that represents a  **DocumentProperty** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
|pfLinkRetVal|Required| **Boolean**|Indicates whether the document property is linked to the container document.|

## Remarks
<a name="sectionSection1"> </a>

This property applies only to custom document properties. For built-in document properties, the value of this property is  **False**.

Use the  **LinkSource**property to set the source for the specified linked property. Setting the  **LinkSource** property sets the **LinkToContent** property to **True**.


## Example
<a name="sectionSection2"> </a>

This example displays the linked status of the custom document property. For the example to work,  **dp** must be a valid **DocumentProperty** object.


```
Sub DisplayLinkStatus(dp As DocumentProperty) 
 Dim stat As String, tf As String 
 If dp.LinkToContent Then 
 tf = "" 
 Else 
 tf = "not " 
 End If 
 stat = "This property is " &amp; tf &amp; "linked" 
 If dp.LinkToContent Then 
 stat = stat + Chr(13) &amp; "The link source is " &amp; dp.LinkSource 
 End If 
 MsgBox stat 
End Sub
```


## See also
<a name="sectionSection2"> </a>


#### Concepts


 [DocumentProperty Object](dd54ca3c-e0e2-4816-539a-17c5b4a928b1.md)
 [Sync Object](1cb049a0-a803-969a-7923-15ddb8da8f3b.md)
#### Other resources


 [DocumentProperty Object Members](568da0ff-fa90-150a-06ec-611de886334e.md)
 [Sync Object Members](748726bd-83de-425a-5af8-177c34e3a013.md)
