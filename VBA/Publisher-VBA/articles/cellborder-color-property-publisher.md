---
title: "Свойство CellBorder.Color (издатель)"
keywords: vbapb10.chm5242882
f1_keywords: vbapb10.chm5242882
ms.prod: publisher
api_name: Publisher.CellBorder.Color
ms.assetid: 59a43522-f0df-fe1a-6e35-19cb012b103f
ms.date: 06/08/2017
ms.openlocfilehash: 8212253c1bd8615277ce14323299631159e8f10d
ms.sourcegitcommit: 1102fd44df64f18dc0561d0b3a7103ca81e74318
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 07/26/2017
---
# <a name="cellbordercolor-property-publisher"></a>Свойство CellBorder.Color (издатель)

Возвращает объект **[ColorFormat](colorformat-object-publisher.md)** , представляющее цвет сведения для указанного объекта.


## <a name="syntax"></a>Синтаксис

 _выражение_. **Цвет**

 переменная _expression_A, представляет собой объект- **CellBorder** .


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


