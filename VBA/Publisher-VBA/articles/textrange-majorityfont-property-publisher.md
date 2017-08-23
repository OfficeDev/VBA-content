---
title: "Свойство TextRange.MajorityFont (издатель)"
keywords: vbapb10.chm5308467
f1_keywords: vbapb10.chm5308467
ms.prod: publisher
api_name: Publisher.TextRange.MajorityFont
ms.assetid: b0007ebc-ed0b-aab8-49fe-76353efbc1d2
ms.date: 06/08/2017
ms.openlocfilehash: 0d830bee09d5f11423370897b0eeec77f20e4886
ms.sourcegitcommit: 1102fd44df64f18dc0561d0b3a7103ca81e74318
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 07/26/2017
---
# <a name="textrangemajorityfont-property-publisher"></a>Свойство TextRange.MajorityFont (издатель)

Возвращает объект **[шрифта](font-object-publisher.md)** , который представляет имя шрифта, используемых в наиболее в диапазон текста.


## <a name="syntax"></a>Синтаксис

 _выражение_. **MajorityFont**

 переменная _expression_A, представляющий объект **TextRange** .


### <a name="return-value"></a>Возвращаемое значение

Font


## <a name="example"></a>Пример

В этом примере создается новое текстовое поле, заполняет его текстом, проверяет, если шрифта, используемых в наиболее Tahoma, а в противном случае изменяет шрифт Tahoma.


```vb
Sub SetFontName() 
 Dim intCount As Integer 
 With ActiveDocument.Pages(1).Shapes _ 
 .AddTextbox(Orientation:=pbTextOrientationHorizontal, _ 
 Left:=100, Top:=100, Width:=100, Height:=100) _ 
 .TextFrame.TextRange 
 For intCount = 1 To 10 
 .InsertAfter NewText:="This is a test. " 
 Next intCount 
 If .MajorityFont <> "Tahoma" Then _ 
 .Font.Name = "Tahoma" 
 End With 
End Sub
```


