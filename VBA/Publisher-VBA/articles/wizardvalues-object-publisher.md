---
title: "Объект WizardValues (издатель)"
keywords: vbapb10.chm1703935
f1_keywords: vbapb10.chm1703935
ms.prod: publisher
api_name: Publisher.WizardValues
ms.assetid: 559659bb-6c9f-9325-c931-14044c059e18
ms.date: 06/08/2017
ms.openlocfilehash: 4938b56a8d994e96c7f399e4ce4e3f1f59bb8896
ms.sourcegitcommit: 1102fd44df64f18dc0561d0b3a7103ca81e74318
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 07/26/2017
---
# <a name="wizardvalues-object-publisher"></a>Объект WizardValues (издатель)

Представляет полный набор допустимых значений для свойства мастера.
 


## <a name="example"></a>Пример

Используйте свойство **[значения](wizardproperty-values-property-publisher.md)** объекта **WizardProperty** для возврата коллекции **WizardValues** . Следующий пример отображает текущее значение для первого свойства мастера в активной публикации и выводит список всех возможных значений.
 

 

```
Dim valAll As WizardValues 
Dim valLoop As WizardValue 
 
With ActiveDocument.Wizard 
 Set valAll = .Properties(1).Values 
 
 MsgBox "Wizard: " &amp; .Name &amp; vbLf &amp; _ 
 "Property: " &amp; .Properties(1).Name &amp; vbLf &amp; _ 
 "Current value: " &amp; .Properties(1).CurrentValueId 
 
 For Each valLoop In valAll 
 MsgBox "Possible value: " &amp; valLoop.ID &amp; " (" &amp; valLoop.Name &amp; ")" 
 Next valLoop 
End With
```


## <a name="properties"></a>Properties



|**Name**|
|:-----|
|[Приложения](wizardvalues-application-property-publisher.md)|
|[Count](wizardvalues-count-property-publisher.md)|
|[Элемент](wizardvalues-item-property-publisher.md)|
|[Родительский раздел](wizardvalues-parent-property-publisher.md)|

