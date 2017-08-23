---
title: "Свойство Document.IsWizard (издатель)"
keywords: vbapb10.chm196745
f1_keywords: vbapb10.chm196745
ms.prod: publisher
api_name: Publisher.Document.IsWizard
ms.assetid: 61ee1a16-eccb-908f-2b34-eee03175c37e
ms.date: 06/08/2017
ms.openlocfilehash: 17394be6cc29438d339ae1286e739cc9bc9893fc
ms.sourcegitcommit: 1102fd44df64f18dc0561d0b3a7103ca81e74318
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 07/26/2017
---
# <a name="documentiswizard-property-publisher"></a>Свойство Document.IsWizard (издатель)

Возвращает **значение True** , если указанная публикация публикации, созданный мастером Microsoft Publisher. Только для чтения **типа Boolean**.


## <a name="syntax"></a>Синтаксис

 _выражение_. **IsWizard**

 переменная _expression_A, представляющий объект **документа** .


### <a name="return-value"></a>Возвращаемое значение

Логический


## <a name="remarks"></a>Заметки

Используйте **[Мастер](document-wizard-property-publisher.md)** свойство объекта **[Document](document-object-publisher.md)** для доступа к мастеру для указанной публикации.


## <a name="example"></a>Пример

Следующий пример проверяет для определения, является ли активным документом мастера публикации. Если он установлен, будут возвращены определенных свойств мастера.


```vb
With ActiveDocument 
 If .IsWizard = True Then 
 Debug.Print .Wizard.Name 
 Debug.Print .Wizard.ID 
 End If 
End With
```


