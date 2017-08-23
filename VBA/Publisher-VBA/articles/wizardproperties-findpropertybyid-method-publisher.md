---
title: "Метод WizardProperties.FindPropertyById (издатель)"
keywords: vbapb10.chm1507332
f1_keywords: vbapb10.chm1507332
ms.prod: publisher
api_name: Publisher.WizardProperties.FindPropertyById
ms.assetid: 9d13ffa2-f251-0e7d-2f36-c747413143d0
ms.date: 06/08/2017
ms.openlocfilehash: a6a066c70d4f2f4aeff970a1842e2596d11a9434
ms.sourcegitcommit: 1102fd44df64f18dc0561d0b3a7103ca81e74318
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 07/26/2017
---
# <a name="wizardpropertiesfindpropertybyid-method-publisher"></a>Метод WizardProperties.FindPropertyById (издатель)

Возвращает объект **[WizardProperty](wizardproperty-object-publisher.md)** , основанный на указанным Идентификатором из коллекции мастер свойства, связанные с публикации проекта или мастер объектов макетов.


## <a name="syntax"></a>Синтаксис

 _выражение_. **FindPropertyById** ( **_Код_**)

 переменная _expression_A, представляет собой объект- **WizardProperties** .


### <a name="parameters"></a>Параметры



|**Имя**|**Обязательный или необязательный**|**Тип данных**|**Описание**|
|:-----|:-----|:-----|:-----|
|ID|Обязательное свойство.| **Длинный**|Идентификатор мастера свойство для возвращения; соответствует свойству **[ID](wizardproperty-id-property-publisher.md)** объекта **WizardProperty** .|

### <a name="return-value"></a>Возвращаемое значение

WizardProperty


## <a name="example"></a>Пример

В следующем примере изменяется параметры макете публикации (информационный бюллетень мастер), чтобы публикация имеет области, выделенной для адреса клиента (адреса клиента).


```vb
Sub SetWizardProperties 
 Dim wizTemp As Wizard 
 Dim wizproTemp As WizardProperty 
 
 Set wizTemp = ActiveDocument.Wizard 
 
 With wizTemp.Properties 
 Set wizproTemp = .FindPropertyById(ID:=901) 
 wizproTemp.CurrentValueId = 1 
 End With 
 
End Sub
```


