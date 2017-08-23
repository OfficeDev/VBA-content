---
title: "Свойство TextStyle.NextParagraphStyle (издатель)"
keywords: vbapb10.chm5963784
f1_keywords: vbapb10.chm5963784
ms.prod: publisher
api_name: Publisher.TextStyle.NextParagraphStyle
ms.assetid: 2b31b883-c26d-3be8-7145-f8e3cf1ba5cc
ms.date: 06/08/2017
ms.openlocfilehash: 0e54bf0635f3beece10a75060fc7c589afd73999
ms.sourcegitcommit: 1102fd44df64f18dc0561d0b3a7103ca81e74318
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 07/26/2017
---
# <a name="textstylenextparagraphstyle-property-publisher"></a>Свойство TextStyle.NextParagraphStyle (издатель)

Возвращает или задает **строку** , представляющую стиль абзаца, стиль указанный текст при нажатии клавиши ВВОД. Чтение и запись.


## <a name="syntax"></a>Синтаксис

 _выражение_. **NextParagraphStyle**

 переменная _expression_A, представляющий объект **стиля текста** .


### <a name="return-value"></a>Возвращаемое значение

String


## <a name="example"></a>Пример

В этом примере создается новый стиль текста и указывает, что следующий новый стиль текста стиля текста — обычный стиль.


```vb
Sub CreateNewTextStyle() 
 Dim styNew As TextStyle 
 Dim fntStyle As Font 
 
 Set styNew = ActiveDocument.TextStyles.Add(StyleName:="Heading 1") 
 Set fntStyle = styNew.Font 
 
 With fntStyle 
 .Name = "Tahoma" 
 .Bold = msoTrue 
 .Size = 15 
 End With 
 
 With styNew 
 .Font = fntStyle 
 .NextParagraphStyle = "Normal" 
 End With 
End Sub
```


