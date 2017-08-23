---
title: "Свойство Document.DocumentDirection (издатель)"
keywords: vbapb10.chm196648
f1_keywords: vbapb10.chm196648
ms.prod: publisher
api_name: Publisher.Document.DocumentDirection
ms.assetid: b28961ad-7adc-3920-0e67-88bb53310d9b
ms.date: 06/08/2017
ms.openlocfilehash: 8c882f51979b6f7ee8d37c478245c338683504ed
ms.sourcegitcommit: 1102fd44df64f18dc0561d0b3a7103ca81e74318
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 07/26/2017
---
# <a name="documentdocumentdirection-property-publisher"></a>Свойство Document.DocumentDirection (издатель)

Возвращает или задает значение константы **PbDirectionType** , которое указывает, доступна ли текст в документ для чтения слева направо или справа налево. Чтение и запись.


## <a name="syntax"></a>Синтаксис

 _выражение_. **DocumentDirection**

 переменная _expression_A, представляющий объект **Document** .


### <a name="return-value"></a>Возвращаемое значение

PbDirectionType


## <a name="remarks"></a>Заметки

Значение свойства **DocumentDirection** может иметь одно из **[PbDirectionType](pbdirectiontype-enumeration-publisher.md)** константы в библиотеке типов, Microsoft Publisher.

Свойство **DocumentDirection** влияет на способ читать документ, но не направление текста в документе. Например если в документе есть пограничный привязки и печати на обеих сторонах страницы, переплета документа справа налево будет отличаться от переплета документа справа налево.

Форматирование направление потока текста, используйте свойство **[DefaultTextFlowDirection](options-defaulttextflowdirection-property-publisher.md)** , чтобы указать направление текста по умолчанию для всего документа или используйте свойство **[Orientation](textframe-orientation-property-publisher.md)** для отдельных текстовой рамки для указания направление текста, используемый по умолчанию, указанный текст только для рамки.


## <a name="example"></a>Пример

В этом примере задается active публикации для чтения слева направо.


```vb
Sub SetBiDiText() 
 ActiveDocument.DocumentDirection = pbDirectionRightToLeft 
End Sub
```


