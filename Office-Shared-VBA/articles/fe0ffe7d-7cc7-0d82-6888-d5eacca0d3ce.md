
# SignatureInfo Object (Office)

 **Last modified:** July 28, 2015

Represents the information used to create a digital or in-document signature.

## Example

The following example uses the  **GetCertificationDetails** method of the **SignatureInfo** object to get the expiration date of the digital certificate.


```
Sub GetCertDetails() 
Dim objSignatureInfo As SignatureInfo 
Dim varDetail As Variant 
 
strDetail = objSignatureInfo.GetCertificationDetail(certdetExpirationDate) 
 
End Sub 

```


## See also


#### Concepts


 [Object Model Reference](499c789a-aba2-0fad-649a-0ea964cd3b5e.md)
#### Other resources


 [SignatureInfo Object Members](52c19097-8afb-d35c-a9f7-eae81e91c05d.md)
