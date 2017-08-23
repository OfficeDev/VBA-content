---
title: "Свойство TextFrame.HasNextLink (издатель)"
keywords: vbapb10.chm3866640
f1_keywords: vbapb10.chm3866640
ms.prod: publisher
api_name: Publisher.TextFrame.HasNextLink
ms.assetid: 907ec470-e283-906a-e25f-f5a8548a18a4
ms.date: 06/08/2017
ms.openlocfilehash: 850c99179b4b6f6ab770cd385c9db59c3c7fa8fc
ms.sourcegitcommit: 1102fd44df64f18dc0561d0b3a7103ca81e74318
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 07/26/2017
---
# <a name="textframehasnextlink-property-publisher"></a>Свойство TextFrame.HasNextLink (издатель)

Указывает, имеет ли frame указанный текст ссылки допустимый прямого текстовое поле. Только для чтения.


## <a name="syntax"></a>Синтаксис

 _выражение_. **HasNextLink**

 переменная _expression_A, представляет собой объект- **TextFrame** .


### <a name="return-value"></a>Возвращаемое значение

MsoTriState


## <a name="remarks"></a>Заметки

Значение свойства **HasNextLink** может иметь одно из **MsoTriState** константы объявляются в библиотеке типов Microsoft Office и показаны в следующей таблице.



|**Константы**|**Описание**|
|:-----|:-----|
| **msoFalse**|Frame указанный текст не имеет ссылки вперед текстовое поле.|
| **msoTrue**| Указанный текст frame имеет ссылки вперед текстовое поле.|

## <a name="example"></a>Пример

Если существует ссылки в этом примере разрывов все ссылки в документе на первый кадр указанный текст. В этом примере предполагает наличие по крайней мере один фигуры на первой странице active публикации.


```vb
Sub AddPreviousNextLinkPages() 
 With ActiveDocument.Pages(1).Shapes(1).TextFrame 
 If .HasNextLink Then .BreakForwardLink 
 If .HasPreviousLink Then .PreviousLinkedTextFrame _ 
 .BreakForwardLink 
 End With 
End Sub
```


