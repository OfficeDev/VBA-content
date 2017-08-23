---
title: "Свойство WizardProperty.CurrentValueId (издатель)"
keywords: vbapb10.chm1572869
f1_keywords: vbapb10.chm1572869
ms.prod: publisher
api_name: Publisher.WizardProperty.CurrentValueId
ms.assetid: d8a2eeb0-f6e7-2687-5952-cddd2cc3914b
ms.date: 06/08/2017
ms.openlocfilehash: 826e21c9823e3dfa4ebab7881f90a41ca340c8c3
ms.sourcegitcommit: 1102fd44df64f18dc0561d0b3a7103ca81e74318
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 07/26/2017
---
# <a name="wizardpropertycurrentvalueid-property-publisher"></a>Свойство WizardProperty.CurrentValueId (издатель)

Возвращает или задает **Long** , указывающее значения параметра в указанной публикации проекта или мастер макетов объектов. Чтение и запись.


## <a name="syntax"></a>Синтаксис

 _выражение_. **CurrentValueId**

 переменная _expression_A, представляет собой объект- **WizardProperty** .


### <a name="return-value"></a>Возвращаемое значение

Длинный


## <a name="remarks"></a>Заметки

Доступ к этому свойству для параметра публикации проекта, для свойства **[Enabled](wizardproperty-enabled-property-publisher.md)** является **False** приводит к ошибке.


## <a name="example"></a>Пример

В следующем примере изменяется параметры макете публикации (информационный бюллетень мастер), чтобы публикация имеет области, выделенной для адреса клиента.


```vb
Dim wizTemp As Wizard 
Dim wizproAll As WizardProperties 
 
Set wizTemp = ActiveDocument.Wizard 
 
With wizTemp.Properties 
 .FindPropertyById(ID:=901).CurrentValueId = 1 
End With
```


