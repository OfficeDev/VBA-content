
# RuleConditions.From Property (Outlook)

 **Last modified:** July 28, 2015

Returns a  ** [ToOrFromRuleCondition](ec5cae2a-cde8-5681-6a49-74e2f0226a4f.md)** object with a ** [ToOrFromRuleCondition.ConditionType](a5c6e08c-643e-965d-cd3e-b434f20579a0.md)** of **olConditionFrom**. Read-only.

## Syntax

 _expression_. **From**

 _expression_A variable that represents a  **RuleConditions** object.


## Remarks

Use the returned  **ToOrFromRuleCondition** object when enumerating the rule conditions or exception conditions of an existing rule, or when creating a new rule that specifies the condition or exception condition that the sender of the message is in the specified list of ** [Recipients](774f56b7-4de8-9584-60cd-4fbf361f4c85.md)**.

This property of the  ** [RuleConditions](e8e9a05a-b36b-add2-b294-8cdc5a97e119.md)** collection always returns a **ToOrFromRuleCondition** object regardless of whether the rule associated with this **RuleConditions** collection has defined such a rule condition. If the rule has defined and enabled such a rule condition, then ** [ToOrFromRuleCondition.Enabled](31e43906-b47a-95e3-d51b-3fa6af553fad.md)** will be **True**.


## See also


#### Concepts


 [RuleConditions Object](e8e9a05a-b36b-add2-b294-8cdc5a97e119.md)
#### Other resources


 [RuleConditions Object Members](b2af6ebf-f9f8-8106-20a3-1725c3b78174.md)
