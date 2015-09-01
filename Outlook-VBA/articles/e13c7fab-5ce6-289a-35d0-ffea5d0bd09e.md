
# SharingItem.BCC Property (Outlook)

 **Last modified:** July 28, 2015

Returns a  **String** representing the display list of blind carbon copy (BCC) names for a ** [SharingItem](63dd3451-44f3-7cc4-c6e2-7dad5835a7d2.md)**. Read/write.

## Syntax

 _expression_. **BCC**

 _expression_A variable that represents a  **SharingItem** object.


## Remarks

This property contains only the display names, delimited with semicolon (;) characters. The  ** [Recipients](774f56b7-4de8-9584-60cd-4fbf361f4c85.md)**collection should be used to modify the BCC recipients. 


 **Note**  If the  **SharingItem** uses an Exchange sharing context, then setting this property to any value other than **Nothing** prevents the item from being sent and causes the ** [Send](54f92175-0e99-f96a-56de-5fc66d97d80f.md)** method to raise an error.


## See also


#### Concepts


 [SharingItem Object](63dd3451-44f3-7cc4-c6e2-7dad5835a7d2.md)
#### Other resources


 [SharingItem Object Members](719ad60e-2242-2c54-778f-006b61690389.md)
