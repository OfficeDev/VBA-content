
# AddressRuleCondition Object (Outlook)

Represents a rule condition that evaluates whether the address for the recipient or sender of the message is contained in the address specified in  **[AddressRuleCondition.Address](de4186ec-0741-8ff6-7789-af0a46c470e0.md)**.


## Remarks

 **AddressRuleCondition** is derived from the **[RuleCondition](e03f91c2-2c08-b036-104a-d6246f28bc2d.md)** object. Each rule is associated with a **[RuleConditions](e8e9a05a-b36b-add2-b294-8cdc5a97e119.md)** object which has a **[RecipientAddress](1b8f361e-0481-75dc-e66e-2bc69228773a.md)** property and a **[SenderAddress](6e5eb1cc-385f-b1b2-aea7-12629cc31030.md)**. Each of these properties always returns a **AddressRuleCondition** object. **[AddressRuleCondition.ConditionType](8b531745-1a4d-d903-5c7d-465b9fd8cbf3.md)** distinguishes among these rule conditions. If the rule has any of these rule conditions enabled, then **[AddressRuleCondition.Enabled](170cd84c-4733-0801-c411-34736e2e1a06.md)** would be **True**.

For more information on specifying rule actions, see [Specifying Rule Conditions](http://msdn.microsoft.com/library/812c131a-fe23-1b8b-5e2d-9459d7102630%28Office.15%29.aspx).


## Properties



|**Name**|
|:-----|
|[Address](de4186ec-0741-8ff6-7789-af0a46c470e0.md)|
|[Application](bc908e8a-83eb-03e7-5b98-9dc0918a67a6.md)|
|[Class](566eb9a5-2b7a-1833-f803-60a750fda257.md)|
|[ConditionType](8b531745-1a4d-d903-5c7d-465b9fd8cbf3.md)|
|[Enabled](170cd84c-4733-0801-c411-34736e2e1a06.md)|
|[Parent](8943ab05-a3c7-6ee2-c2c1-f97315a08ac0.md)|
|[Session](c5134be6-7ce4-dc65-8bde-9c725ef3ba8c.md)|

## See also


#### Other resources


[Outlook Object Model Reference](http://msdn.microsoft.com/library/73221b13-d8d8-99b8-3394-b95dbbfd5ddc%28Office.15%29.aspx)