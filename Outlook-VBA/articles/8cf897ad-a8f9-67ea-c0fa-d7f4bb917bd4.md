
# AddressRuleCondition Object (Outlook)

 **Last modified:** July 28, 2015

Represents a rule condition that evaluates whether the address for the recipient or sender of the message is contained in the address specified in  ** [AddressRuleCondition.Address](de4186ec-0741-8ff6-7789-af0a46c470e0.md)**.

## Remarks

 **AddressRuleCondition** is derived from the ** [RuleCondition](e03f91c2-2c08-b036-104a-d6246f28bc2d.md)** object. Each rule is associated with a ** [RuleConditions](e8e9a05a-b36b-add2-b294-8cdc5a97e119.md)** object which has a ** [RecipientAddress](1b8f361e-0481-75dc-e66e-2bc69228773a.md)** property and a ** [SenderAddress](6e5eb1cc-385f-b1b2-aea7-12629cc31030.md)**. Each of these properties always returns a  **AddressRuleCondition** object. ** [AddressRuleCondition.ConditionType](8b531745-1a4d-d903-5c7d-465b9fd8cbf3.md)** distinguishes among these rule conditions. If the rule has any of these rule conditions enabled, then ** [AddressRuleCondition.Enabled](170cd84c-4733-0801-c411-34736e2e1a06.md)** would be **True**.

For more information on specifying rule actions, see  [Specifying Rule Conditions](812c131a-fe23-1b8b-5e2d-9459d7102630.md).


## See also


#### Concepts


 [Outlook Object Model Reference](73221b13-d8d8-99b8-3394-b95dbbfd5ddc.md)
#### Other resources


 [AddressRuleCondition Object Members](d15b0554-6b47-b201-fd41-744ea056d3f6.md)
