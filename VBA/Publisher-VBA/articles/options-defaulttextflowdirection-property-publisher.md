---
title: "Свойство Options.DefaultTextFlowDirection (издатель)"
keywords: vbapb10.chm1048628
f1_keywords: vbapb10.chm1048628
ms.prod: publisher
api_name: Publisher.Options.DefaultTextFlowDirection
ms.assetid: 7c17768a-cd9c-704d-fa27-f0dfd7648054
ms.date: 06/08/2017
ms.openlocfilehash: cc6a44d312904d0213133ffbc67d7a7f94767141
ms.sourcegitcommit: 1102fd44df64f18dc0561d0b3a7103ca81e74318
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 07/26/2017
---
# <a name="optionsdefaulttextflowdirection-property-publisher"></a>Свойство Options.DefaultTextFlowDirection (издатель)

Возвращает или задает значение константы **PbDirectionType** , представляющий глобальный параметр Microsoft Publisher, указывающее, передается ли текста слева направо или справа налево в публикации. Чтение и запись.


## <a name="syntax"></a>Синтаксис

 _выражение_. **DefaultTextFlowDirection**

 переменная _expression_A, представляет собой объект- **Параметры** .


### <a name="return-value"></a>Возвращаемое значение

PbDirectionType


## <a name="remarks"></a>Заметки

Значение свойства **DefaultTextFlowDirection** может иметь одно из **[PbDirectionType](pbdirectiontype-enumeration-publisher.md)** константы в библиотеке типов, Publisher.

Это свойство приводит к ошибке, если не выполняется в бизнес-аналитики обмен сообщениями версии Publisher (например, арабский).


## <a name="example"></a>Пример

В этом примере направление текста по умолчанию для новых публикаций и текста в бизнес-аналитики обмен сообщениями версии Publisher.


```vb
Sub SetDefaultDirection() 
 With Options 
 .DefaultPubDirection = pbDirectionRightToLeft 
 .DefaultTextFlowDirection = pbDirectionRightToLeft 
 End With 
End Sub
```


