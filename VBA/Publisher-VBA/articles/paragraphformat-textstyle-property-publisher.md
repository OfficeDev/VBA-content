---
title: "Свойство ParagraphFormat.TextStyle (издатель)"
keywords: vbapb10.chm5439508
f1_keywords: vbapb10.chm5439508
ms.prod: publisher
api_name: Publisher.ParagraphFormat.TextStyle
ms.assetid: 8495c9c8-387e-a2e8-26cb-08f660dde985
ms.date: 06/08/2017
ms.openlocfilehash: 565045948fe40bffb7d167314fe223aa2c5892b3
ms.sourcegitcommit: 1102fd44df64f18dc0561d0b3a7103ca81e74318
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 07/26/2017
---
# <a name="paragraphformattextstyle-property-publisher"></a>Свойство ParagraphFormat.TextStyle (издатель)

Возвращает или задает **Variant** , который представляет текст стиль абзаца. Чтение и запись.


## <a name="syntax"></a>Синтаксис

 _выражение_. **Стиля текста**

 переменная _expression_A, представляет собой объект- **ParagraphFormat** .


### <a name="return-value"></a>Возвращаемое значение

Variant


## <a name="example"></a>Пример

В этом примере изменяется стиль текста текущего выбора, если выделение не будет отформатирован стиль Обычный текст. В этом примере предполагается, что в активной публикации выбранного текста.


```vb
Sub SetTextStyle() 
 With Selection.TextRange.ParagraphFormat 
 If .TextStyle <> "Normal" Then _ 
 .TextStyle = "Normal" 
 End With 
End Sub
```


