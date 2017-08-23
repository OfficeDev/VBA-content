---
title: "Свойство Tag.Value (издатель)"
keywords: vbapb10.chm4718596
f1_keywords: vbapb10.chm4718596
ms.prod: publisher
api_name: Publisher.Tag.Value
ms.assetid: dee3b69b-ae5b-df13-561e-84105057979a
ms.date: 06/08/2017
ms.openlocfilehash: f2b0b8bb38bbaae8ecf5d8788153be54533b07b6
ms.sourcegitcommit: 1102fd44df64f18dc0561d0b3a7103ca81e74318
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 07/26/2017
---
# <a name="tagvalue-property-publisher"></a>Свойство Tag.Value (издатель)

Возвращает или задает **Variant** , который представляет значение тега фигуры, страницы или публикации. Чтение и запись.


## <a name="syntax"></a>Синтаксис

 _выражение_. **Значение**

 переменная _expression_A, представляет собой объект- **тег** .


## <a name="example"></a>Пример

В этом примере создается новый тег для активной публикации и затем отображает значение тега.


```vb
Sub CreatePublicationTag() 
 With ActiveDocument 
 .Tags.Add Name:="ActivePub", Value:="This is the active publication." 
 MsgBox .Tags(1).Value 
 End With 
End Sub
```


