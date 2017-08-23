---
title: "Объект WizardValue (издатель)"
keywords: vbapb10.chm2162687
f1_keywords: vbapb10.chm2162687
ms.prod: publisher
api_name: Publisher.WizardValue
ms.assetid: 15b60632-d1b1-c62b-0264-72d65bd1fe82
ms.date: 06/08/2017
ms.openlocfilehash: f2399e8c108052bde7763acd79d1fedcf4f96c92
ms.sourcegitcommit: 1102fd44df64f18dc0561d0b3a7103ca81e74318
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 07/26/2017
---
# <a name="wizardvalue-object-publisher"></a>Объект WizardValue (издатель)

Представляет допустимое значение для свойства указанного мастера.
 


## <a name="example"></a>Пример

Свойство **[Item](wizardvalues-item-property-publisher.md)** коллекции **WizardValues** возвращает объект **WizardValue** . Следующий пример отображает текущее значение для первого свойства мастера в активной публикации и выводит список всех возможных значений.
 

 

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
|[Приложения](wizardvalue-application-property-publisher.md)|
|[ID](wizardvalue-id-property-publisher.md)|
|[Name](wizardvalue-name-property-publisher.md)|
|[Родительский раздел](wizardvalue-parent-property-publisher.md)|

