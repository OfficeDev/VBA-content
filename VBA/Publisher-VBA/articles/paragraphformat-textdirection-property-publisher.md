---
title: "Свойство ParagraphFormat.TextDirection (издатель)"
keywords: vbapb10.chm5439507
f1_keywords: vbapb10.chm5439507
ms.prod: publisher
api_name: Publisher.ParagraphFormat.TextDirection
ms.assetid: b96c634d-0e7e-dba8-2bf4-e5baf3afa3d1
ms.date: 06/08/2017
ms.openlocfilehash: 9bb79cb509ca854746457e653d9f93f4021c38f0
ms.sourcegitcommit: 1102fd44df64f18dc0561d0b3a7103ca81e74318
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 07/26/2017
---
# <a name="paragraphformattextdirection-property-publisher"></a>Свойство ParagraphFormat.TextDirection (издатель)

Возвращает или задает значение, указывающее направление, в какой текст располагается в указанном абзаце константы **PbTextDirection** . Чтение и запись.


## <a name="syntax"></a>Синтаксис

 _выражение_. **TextDirection**

 переменная _expression_A, представляет собой объект- **ParagraphFormat** .


### <a name="return-value"></a>Возвращаемое значение

PbTextDirection


## <a name="remarks"></a>Заметки

Данное свойство предназначен для использования совместно с документами, содержащими текст на языках слева направо и справа налево. Установка для свойства значение, не соответствует направление текста, зависит от используемого языка может привести к непредсказуемым результатам.

Значение свойства **TextDirection** может быть одной из констант **PbTextDirection** объявлена в библиотеке типов, Microsoft Publisher и показаны в следующей таблице.



|**Константы**|**Описание**|
|:-----|:-----|
| **pbTextDirectionLeftToRight**| Текст располагается слева направо.|
| **pbTextDirectionMixed**|Возвращает значение, указывающее, диапазон, содержащий фрагменту текста слева направо и фрагменту текста справа налево.|
| **pbTextDirectionRightToLeft**|Потоки текста справа налево.|

## <a name="example"></a>Пример

В следующем примере изменяется направление текста первой фигуры на страницу, чтобы он потоков для письма справа налево.


```vb
ActiveDocument.Pages(1).Shapes(1).TextFrame.TextRange _ 
 .ParagraphFormat.TextDirection = pbTextDirectionRightToLeft
```


