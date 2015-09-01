
# FromRssFeedRuleCondition.FromRssFeed Property (Outlook)

 **Last modified:** July 28, 2015

Returns or sets an array of  **String** elements that represent the RSS subscriptions that are being evaluated by the rule condition. Read/write.

## Syntax

 _expression_. **FromRssFeed**

 _expression_A variable that represents a  **FromRssFeedRuleCondition** object.


## Remarks

Each element of the array is a single RSS subscription. Multiple RSS subscriptions are evaluated as logical OR conditions.

You cannot obtain a list of valid RSS subscription names through the Outlook object model. You can obtain a list of valid RSS subscription names from the XML file Outlook.Sharing.xml.obi, which is located in the folder [drive]\Documents and Settings\[UserName]\Local Settings\Application Data\Microsoft\Outlook\. The  `name` attribute of the <local> tag contains the name of the RSS subscription that must be supplied in the array of strings for **FromRssFeed**. To enumerate all RSS subscriptions, examine the <bindings> tag where  `<binding prov="{0006F0AF-0000-0000-C000-000000000046}">`.

Returns an error if one or more elements in the array contains an empty string or an invalid RSS subscription.


## See also


#### Concepts


 [FromRssFeedRuleCondition Object](8de6e629-7e3d-b4df-d758-a5bff3abd6a1.md)
#### Other resources


 [FromRssFeedRuleCondition Object Members](0c0a949a-d654-6701-f70d-9a5bb908fed8.md)
