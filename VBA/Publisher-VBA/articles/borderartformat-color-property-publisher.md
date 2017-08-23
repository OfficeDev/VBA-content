---
title: "Свойство BorderArtFormat.Color (издатель)"
keywords: vbapb10.chm7602183
f1_keywords: vbapb10.chm7602183
ms.prod: publisher
api_name: Publisher.BorderArtFormat.Color
ms.assetid: fb2fe2f7-d321-43d3-232d-db3b513dae43
ms.date: 06/08/2017
ms.openlocfilehash: 7ff1877014959f205aaa9342c75996adf66d2869
ms.sourcegitcommit: 1102fd44df64f18dc0561d0b3a7103ca81e74318
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 07/26/2017
---
# <a name="borderartformatcolor-property-publisher"></a>Свойство BorderArtFormat.Color (издатель)

Возвращает объект **[ColorFormat](colorformat-object-publisher.md)** , представляющее цвет сведения для указанного объекта.


## <a name="syntax"></a>Синтаксис

 _выражение_. **Цвет**

 переменная _expression_A, представляет собой объект- **BorderArtFormat** .


## <a name="example"></a>Пример

В этом примере проверяется цвет шрифта для первой статьи в активном документе и сообщает пользователю, если или не установлен черный цвет шрифта.


```vb
Sub FontColor() 
 
 If Application.ActiveDocument.Stories(1) _ 
 .TextRange.Font.Color.RGB = RGB(Red:=0, Green:=0, Blue:=0) Then 
 MsgBox "Your font color is black" 
 Else 
 MsgBox "Your font color is not black" 
 End If 
 
End Sub
```


## <a name="see-also"></a>См. также


#### <a name="concepts"></a>Основные понятия


 [Объект BorderArtFormat](borderartformat-object-publisher.md)

