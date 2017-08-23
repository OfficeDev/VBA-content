---
title: "Свойство TextFrame.HasPreviousLink (издатель)"
keywords: vbapb10.chm3866641
f1_keywords: vbapb10.chm3866641
ms.prod: publisher
api_name: Publisher.TextFrame.HasPreviousLink
ms.assetid: 85e0b497-55c9-d49f-2b65-e199361c121a
ms.date: 06/08/2017
ms.openlocfilehash: 77b191b93e9f699cc931360a598dbdf99cf79451
ms.sourcegitcommit: 1102fd44df64f18dc0561d0b3a7103ca81e74318
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 07/26/2017
---
# <a name="textframehaspreviouslink-property-publisher"></a>Свойство TextFrame.HasPreviousLink (издатель)

Возвращает **msoTrue** , если frame указанный текст имеет допустимый ссылка на обратной текстовое поле и **msoFalse** , если это не. Только для чтения.


## <a name="syntax"></a>Синтаксис

 _выражение_. **HasPreviousLink**

 переменная _expression_A, представляет собой объект- **TextFrame** .


### <a name="return-value"></a>Возвращаемое значение

MsoTriState


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


