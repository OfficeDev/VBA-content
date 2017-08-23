---
title: "Свойство WizardProperty.Values (издатель)"
keywords: vbapb10.chm1572872
f1_keywords: vbapb10.chm1572872
ms.prod: publisher
api_name: Publisher.WizardProperty.Values
ms.assetid: 478d3b98-65f4-c448-8096-3e999c865846
ms.date: 06/08/2017
ms.openlocfilehash: e855e4873f0c7eb82d8276a00b4c97ceb9da191e
ms.sourcegitcommit: 1102fd44df64f18dc0561d0b3a7103ca81e74318
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 07/26/2017
---
# <a name="wizardpropertyvalues-property-publisher"></a>Свойство WizardProperty.Values (издатель)

Возвращает коллекцию **[WizardValues](wizardvalues-object-publisher.md)** , представляющую все допустимые значения для свойства мастера.


## <a name="syntax"></a>Синтаксис

 _выражение_. **Значения**

 переменная _expression_A, представляет собой объект- **WizardProperty** .


### <a name="return-value"></a>Возвращаемое значение

WizardValues


## <a name="example"></a>Пример

Следующий пример отображает текущее значение для первого свойства мастера в активной публикации и выводит список всех возможных значений.


```vb
Dim valAll As WizardValues 
Dim valLoop As WizardValue 
 
With ActiveDocument.Wizard 
 Set valAll = .Properties(1).Values 
 
 MsgBox "Wizard: " &; .Name &; vbLf &; _ 
 "Property: " &; .Properties(1).Name &; vbLf &; _ 
 "Current value: " &; .Properties(1).CurrentValueId 
 
 For Each valLoop In valAll 
 MsgBox "Possible value: " &; valLoop.ID &; " (" &; valLoop.Name &; ")" 
 Next valLoop 
End With 

```


