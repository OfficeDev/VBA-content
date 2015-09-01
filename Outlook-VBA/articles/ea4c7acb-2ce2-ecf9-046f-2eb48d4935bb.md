
# RuleActions Members (Outlook)
The  **RuleActions** object contains a set of ** [RuleAction](6451788f-e5ed-239c-a34d-b564b52d8955.md)** objects or objects derived from **RuleAction**, representing the actions that are executed on a  ** [Rule](ea2ddbcc-fd65-a636-c6da-79950033f385.md)** object.

 **Last modified:** July 28, 2015


## Methods



|**Name**|**Description**|
|:-----|:-----|
| [Item](d37a3f0c-0273-e4c2-21e5-661484244671.md)|Obtains a  ** [RuleAction](6451788f-e5ed-239c-a34d-b564b52d8955.md)** object specified by _Index_ which is a numerical index into the ** [RuleActions](82ba76cd-86a4-3372-cb51-2df1d58c8b71.md)** collection.|

## Properties



|**Name**|**Description**|
|:-----|:-----|
| [Application](001f7719-084b-2b80-6660-097b5a47c433.md)|Returns an  ** [Application](797003e7-ecd1-eccb-eaaf-32d6ddde8348.md)**object that represents the parent Outlook application for the object. Read-only.|
| [AssignToCategory](7780487b-3dd4-6143-2250-2109872b6192.md)|Returns an  ** [AssignToCategoryRuleAction](402f4742-72ba-2559-4e4c-e2b8248cd7f6.md)** object with ** [AssignToCategoryRuleAction.ActionType](bef50a28-967e-7336-ef0b-2e8edb843c0d.md)** being **olRuleAssignToCategory**. Read-only.|
| [CC](edbaaf74-cfd2-304b-61f3-8d12a621239c.md)|Returns a  ** [SendRuleAction](4ea8f519-8bb3-b0bf-9742-8a492e7ffff7.md)** object with ** [SendRuleAction.ActionType](07b46194-32b4-f04f-d18e-d4b7f3db8f07.md)** being **olRuleActionCcMessage**. Read-only.|
| [Class](99e959aa-7081-aca3-7415-827c6bc3bf64.md)|Returns an  ** [OlObjectClass](33d724b3-df3c-2a7f-a80f-93b66d96f588.md)** constant indicating the object's class. Read-only.|
| [ClearCategories](db594b52-1700-67a7-8445-338f7df254e9.md)|Returns a  ** [RuleAction](6451788f-e5ed-239c-a34d-b564b52d8955.md)** object with a ** [RuleAction.ActionType](5701cd66-2f45-ae24-12b8-fc5e27bf8742.md)** of **olRuleActionClearCategories**. Read-only.|
| [CopyToFolder](6e5c0ea8-6287-2904-c8d8-b3c6b5f7cb24.md)|Returns a  ** [MoveOrCopyRuleAction](db951ad8-0d05-1696-acf4-c1da4fbdee33.md)** object with ** [MoveOrCopyRuleAction.ActionType](204bef7d-a19a-abd1-d494-23c33aa9f145.md)** being **olRuleActionCopyToFolder**. Read-only.|
| [Count](91b4425f-0e17-fff1-0d9c-1697b205ff2a.md)|Returns a  **Long** indicating the count of objects in the specified collection. Read-only.|
| [Delete](eb219d46-64c2-650c-ad39-e049ef33160f.md)|Returns a  ** [RuleAction](6451788f-e5ed-239c-a34d-b564b52d8955.md)** object with ** [RuleAction.ActionType](5701cd66-2f45-ae24-12b8-fc5e27bf8742.md)** being **olRuleActionDelete**. Read-only.|
| [DeletePermanently](fbd19516-c599-b7e6-cdd3-0c182d32b216.md)|Returns a  ** [RuleAction](6451788f-e5ed-239c-a34d-b564b52d8955.md)** object with ** [RuleAction.ActionType](5701cd66-2f45-ae24-12b8-fc5e27bf8742.md)** being **olRuleActionDeletePermanently**. Read-only.|
| [DesktopAlert](700c3e5a-ebb1-3cfe-e27d-eea305c27143.md)|Returns a  ** [RuleAction](6451788f-e5ed-239c-a34d-b564b52d8955.md)** object with ** [RuleAction.ActionType](5701cd66-2f45-ae24-12b8-fc5e27bf8742.md)** being **olRuleActionDesktopAlert**. Read-only.|
| [Forward](48315808-5ef7-3262-a305-5b659986e7a8.md)|Returns a  ** [SendRuleAction](4ea8f519-8bb3-b0bf-9742-8a492e7ffff7.md)** object with ** [SendRuleAction.ActionType](07b46194-32b4-f04f-d18e-d4b7f3db8f07.md)** being **olRuleActionForward**. Read-only.|
| [ForwardAsAttachment](9e2eb736-35d9-b17e-8d6d-c5105388f513.md)|Returns a  ** [SendRuleAction](4ea8f519-8bb3-b0bf-9742-8a492e7ffff7.md)** object with ** [SendRuleAction.ActionType](07b46194-32b4-f04f-d18e-d4b7f3db8f07.md)** being **olRuleActionForwardAsAttachment**. Read-only.|
| [MarkAsTask](9dd48e1a-d780-0923-11b0-e980c1fe19ab.md)|Returns a  ** [MarkAsTaskRuleAction](639d9242-7387-2b25-9d0f-f7a14cf16790.md)** object with ** [MarkAsTaskRuleAction.ActionType](d05f10cb-5c5d-37e5-1d6e-a5e4147bd1b3.md)** being **olRuleActionMarkAsTask**. Read-only.|
| [MoveToFolder](6d9c577d-e022-72fc-45f2-bdda7a8761de.md)|Returns a  ** [MoveOrCopyRuleAction](db951ad8-0d05-1696-acf4-c1da4fbdee33.md)** object with ** [MoveOrCopyRuleAction.ActionType](204bef7d-a19a-abd1-d494-23c33aa9f145.md)** being **olRuleActionMoveToFolder**. Read-only.|
| [NewItemAlert](01de8523-7617-c3df-39c6-395f85eda57f.md)|Returns a  ** [NewItemAlertRuleAction](01d30816-50aa-ff23-69a0-4aa627b3d7e4.md)** object with ** [ActionType](e6cb9c35-48c3-f7fe-cfed-4eb45cb83149.md)** being **olRuleActionNewItemAlert**. Read-only.|
| [NotifyDelivery](fd5e3831-6181-8452-10e5-17ff48d54ba7.md)|Returns a  ** [RuleAction](6451788f-e5ed-239c-a34d-b564b52d8955.md)** object with ** [RuleAction.ActionType](5701cd66-2f45-ae24-12b8-fc5e27bf8742.md)** being **olRuleActionNotifyDelivery**. Read-only.|
| [NotifyRead](922a1ea7-8992-0387-e4e1-2e74d6a2cf2a.md)|Returns a  ** [RuleAction](6451788f-e5ed-239c-a34d-b564b52d8955.md)** object with ** [RuleAction.ActionType](5701cd66-2f45-ae24-12b8-fc5e27bf8742.md)** being **olRuleActionNotifyRead**. Read-only.|
| [Parent](697b3625-f184-b6d7-9788-bf74377d636b.md)|Returns the parent  **Object** of the specified object. Read-only.|
| [PlaySound](43a79f2d-9e7b-7053-6901-40e815220ac0.md)|Returns a  ** [PlaySoundRuleAction](6a7a1f78-640e-8ffc-558c-c26b87638d64.md)** object with ** [PlaySoundRuleAction.ActionType](f3b2ec1d-9b8b-64b8-cc02-9d1aec8fd764.md)** being **olRuleActionNotifyRead**. Read-only.|
| [Redirect](a8e13e82-43c5-168a-0334-386fd02489f8.md)|Returns a  ** [SendRuleAction](4ea8f519-8bb3-b0bf-9742-8a492e7ffff7.md)** object with ** [SendRuleAction.ActionType](07b46194-32b4-f04f-d18e-d4b7f3db8f07.md)** being **olRuleActionRedirect**. Read-only.|
| [Session](10b906a5-421c-e858-f8f1-561818425f0a.md)|Returns the  ** [NameSpace](f0dcaa19-07f5-5d42-a3bf-2e42b7885644.md)**object for the current session. Read-only.|
| [Stop](62157e66-dc87-b12e-444d-864d34f4211f.md)|Returns a  ** [RuleAction](6451788f-e5ed-239c-a34d-b564b52d8955.md)** object with ** [RuleAction.ActionType](5701cd66-2f45-ae24-12b8-fc5e27bf8742.md)** being **olRuleActionStop**. Read-only.|
