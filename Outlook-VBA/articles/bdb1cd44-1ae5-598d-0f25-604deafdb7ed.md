
# ContactItem.MailingAddressPostalCode Property (Outlook)

 **Last modified:** July 28, 2015

Returns or sets a  **String** representing the postal code (zip code) portion of the selected mailing address of the contact. Read/write.

## Syntax

 _expression_. **MailingAddressPostalCode**

 _expression_A variable that represents a  **ContactItem** object.


## Remarks

This property replicates the property indicated by the  ** [SelectedMailingAddress](7f0a68a0-2663-276f-7217-f580d63edb51.md)**property, which is one of the following  **OlMailingAddress** constants: **olBusiness**,  **olHome**,  **olNone**, or  **olOther**. While it can be changed or entered independently, any such changes or entries to this property will be overwritten by any subsequent changes or entries to the property indicated by  **SelectedMailingAddress**.


## See also


#### Concepts


 [ContactItem Object](8e32093c-a678-f1fd-3f35-c2d8994d166f.md)
#### Other resources


 [ContactItem Object Members](a8b13369-4c87-02aa-e62a-1f3067e559fa.md)
