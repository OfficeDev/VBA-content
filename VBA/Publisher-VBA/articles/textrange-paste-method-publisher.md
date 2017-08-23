---
title: "Метод TextRange.Paste (издатель)"
keywords: vbapb10.chm5308482
f1_keywords: vbapb10.chm5308482
ms.prod: publisher
api_name: Publisher.TextRange.Paste
ms.assetid: dd29c9ab-7f56-3604-3390-8f5a3b97821f
ms.date: 06/08/2017
ms.openlocfilehash: 22a914831fdc14b20f3dcaf1f7b9f8c9d7957da2
ms.sourcegitcommit: 1102fd44df64f18dc0561d0b3a7103ca81e74318
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 07/26/2017
---
# <a name="textrangepaste-method-publisher"></a>Метод TextRange.Paste (издатель)

Вставляет этот текст в буфер обмена в диапазон указанный текст и возвращает объект **[TextRange](textrange-object-publisher.md)** , представляющий вставленного текста.


## <a name="syntax"></a>Синтаксис

 _выражение_. **Вставить**

 переменная _expression_A, представляющий объект **TextRange** .


### <a name="return-value"></a>Возвращаемое значение

TextRange


## <a name="example"></a>Пример

Этот пример удаляет текст в фигуры на страницу один активный публикации помещает его в буфер обмена и вставляет его после первого слова в двух фигуры на той же странице. В этом примере предполагается, что каждой фигуры содержит текст.


```vb
With ActiveDocument.Pages(1) 
 .Shapes(1).TextFrame.TextRange.Cut 
 .Shapes(2).TextFrame.TextRange. _ 
 Words(1).Paste 
End With 

```


