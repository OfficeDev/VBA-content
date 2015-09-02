
# Speech.Direction Property (Excel)

 **Last modified:** July 28, 2015

 **In this article**
 [Syntax](#sectionSection0)
 [Remarks](#sectionSection1)
 [Example](#sectionSection2)


Returns or sets the order in which the cells will be spoken. The value of the  **Direction** property is an ** [XlSpeakDirection](6e738db7-9722-21ee-5904-1289f9e3987b.md)**constant. Read/write.


## Syntax
<a name="sectionSection0"> </a>

 _expression_. **Direction**

 _expression_A variable that represents a  **Speech** object.


## Remarks
<a name="sectionSection1"> </a>





| **XlSpeakDirection** can be one of these **XlSpeakDirection** constants.|
| **xlSpeakByColumns**|
| **xlSpeakByRows**|

## Example
<a name="sectionSection2"> </a>

In this example, Microsoft Excel determines the speech direction and notifies the user.


```
Sub CheckSpeechDirection() 
 
 ' Notify user of speech direction. 
 If Application.Speech.Direction = xlSpeakByColumns Then 
 MsgBox "The speech direction is set to speak by columns." 
 Else 
 MsgBox "The speech direction is set to speak by rows." 
 End If 
 
End Sub
```


## See also
<a name="sectionSection2"> </a>


#### Concepts


 [Databar Object](2684e913-c278-e6be-ba9d-053b6ad58bae.md)
 [Speech Object](1ddd61bc-989e-4766-423e-515ec5d1c23a.md)
#### Other resources


 [Databar Object Members](137f7e88-bb61-48a3-d2cb-76a8282cd62e.md)
 [Speech Object Members](5dcc198f-153f-0049-d870-bf162cbde9c7.md)
