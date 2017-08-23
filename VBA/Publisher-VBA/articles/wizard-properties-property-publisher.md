---
title: "Свойство Wizard.Properties (издатель)"
keywords: vbapb10.chm1441797
f1_keywords: vbapb10.chm1441797
ms.prod: publisher
api_name: Publisher.Wizard.Properties
ms.assetid: 9f9811b3-10ee-d429-c5a2-8223349525f2
ms.date: 06/08/2017
ms.openlocfilehash: c31905102d58f0bc995644079214b831f8367e05
ms.sourcegitcommit: 1102fd44df64f18dc0561d0b3a7103ca81e74318
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 07/26/2017
---
# <a name="wizardproperties-property-publisher"></a>Свойство Wizard.Properties (издатель)

Возвращает коллекцию **[WizardProperties](wizardproperties-object-publisher.md)** , представляющую все параметры, которые являются частью указанной публикации проекта или мастер макетов объектов.


## <a name="syntax"></a>Синтаксис

 _выражение_. **Свойства**

 переменная _expression_A, представляющий объект **мастера** .


### <a name="return-value"></a>Возвращаемое значение

WizardProperties


## <a name="example"></a>Пример

Следующий пример отчетов по публикации проекта, связанного с активной публикации, отображение его имя и текущие настройки.


```vb
Dim wizTemp As Wizard 
Dim wizproTemp As WizardProperty 
Dim wizproAll As WizardProperties 
 
Set wizTemp = ActiveDocument.Wizard 
 
With wizTemp 
 Set wizproAll = .Properties 
 Debug.Print "Publication Design associated with " _ 
 &; "current publication: " _ 
 &; .Name 
 For Each wizproTemp In wizproAll 
 With wizproTemp 
 Debug.Print " Wizard property: " _ 
 &; .Name &; " = " &; .CurrentValueId 
 End With 
 Next wizproTemp 
End With
```


