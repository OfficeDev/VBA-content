
# ExchangeDistributionList.Address Property (Outlook)

 **Last modified:** July 28, 2015

Returns or sets a  **String** representing the X400 e-mail address of the ** [ExchangeDistributionList](2830dfba-6c0a-a81f-6b98-92ac2aafb59d.md)**. Read/write.

## Syntax

 _expression_. **Address**

 _expression_A variable that represents an  **ExchangeDistributionList** object.


## Remarks

This property assumes the X400 address of the distribution list. To determine the primary Internet address, use the  ** [ExchangeDistributionList.PrimarySmtpAddress](f64bbc29-14c4-be68-402a-16d9ac34a727.md)** property.

The  **Address** property must be set before calling the ** [ExchangeDistributionList.Details](6c93a583-cc61-e527-7832-88dba525854a.md)** method.


## See also


#### Concepts


 [ExchangeDistributionList Object](2830dfba-6c0a-a81f-6b98-92ac2aafb59d.md)
#### Other resources


 [ExchangeDistributionList Object Members](89105487-3e5b-ee8b-02e0-33ad42bd2fbe.md)
