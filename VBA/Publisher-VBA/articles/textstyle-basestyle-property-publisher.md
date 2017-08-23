---
title: "Свойство TextStyle.BaseStyle (издатель)"
keywords: vbapb10.chm5963783
f1_keywords: vbapb10.chm5963783
ms.prod: publisher
api_name: Publisher.TextStyle.BaseStyle
ms.assetid: c8d1665c-c232-ecdf-3c1c-f614c7374c1e
ms.date: 06/08/2017
ms.openlocfilehash: 561b8231cad1df757afd94ca2f114f3719851157
ms.sourcegitcommit: 1102fd44df64f18dc0561d0b3a7103ca81e74318
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 07/26/2017
---
# <a name="textstylebasestyle-property-publisher"></a>Свойство TextStyle.BaseStyle (издатель)

Возвращает или задает **строку** , представляющую стиль, на котором основан форматирование другого стиля. Чтение и запись.


## <a name="syntax"></a>Синтаксис

 _выражение_. **BaseStyle**

 переменная _expression_A, представляющий объект **стиля текста** .


### <a name="return-value"></a>Возвращаемое значение

String


## <a name="example"></a>Пример

В этом примере задается базового форматирования с именем основной текст для форматирования стиля Обычный стиль.


```vb
Sub SetBaseStyle() 
 With ActiveDocument.TextStyles 
 .Add "Body Text" 
 .Item("Body Text").BaseStyle = "Normal" 
 End With 
End Sub
```


