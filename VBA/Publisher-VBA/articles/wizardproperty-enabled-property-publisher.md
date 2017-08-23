---
title: "Свойство WizardProperty.Enabled (издатель)"
keywords: vbapb10.chm1572871
f1_keywords: vbapb10.chm1572871
ms.prod: publisher
api_name: Publisher.WizardProperty.Enabled
ms.assetid: c66741c8-1493-ac90-4ecb-ed8d58743c69
ms.date: 06/08/2017
ms.openlocfilehash: cd1c45264414e75559da8bc065abb044c1282fc8
ms.sourcegitcommit: 1102fd44df64f18dc0561d0b3a7103ca81e74318
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 07/26/2017
---
# <a name="wizardpropertyenabled-property-publisher"></a>Свойство WizardProperty.Enabled (издатель)

 **Значение true,** Если включено свойство мастера. Только для чтения **типа Boolean**.


## <a name="syntax"></a>Синтаксис

 _выражение_. **Включено**

 переменная _expression_A, представляющий объект **WizardProperty** .


### <a name="return-value"></a>Возвращаемое значение

Логический


## <a name="example"></a>Пример

В этом примере имя каждого свойства enabled мастера в активной публикации.


```vb
Sub SetEnabledProperty() 
 Dim wizProperty As WizardProperty 
 For Each wizProperty In ActiveDocument.Wizard.Properties 
 If wizProperty.Enabled = True Then 
 MsgBox "The name of the wizard property is " &; wizProperty.Name 
 End If 
 Next 
End Sub
```


