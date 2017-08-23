---
title: "Свойство Plate.Color (издатель)"
keywords: vbapb10.chm2883587
f1_keywords: vbapb10.chm2883587
ms.prod: publisher
api_name: Publisher.Plate.Color
ms.assetid: 4c4897f5-90bb-cb84-e9b8-47df1a912916
ms.date: 06/08/2017
ms.openlocfilehash: b62085b6be43dd860f5bce26470da493d9a9a017
ms.sourcegitcommit: 1102fd44df64f18dc0561d0b3a7103ca81e74318
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 07/26/2017
---
# <a name="platecolor-property-publisher"></a>Свойство Plate.Color (издатель)

Возвращает объект **[ColorFormat](colorformat-object-publisher.md)** , представляющее цвет сведения для указанного объекта.


## <a name="syntax"></a>Синтаксис

 _выражение_. **Цвет**

 переменная _expression_A, представляющий объект **формы** .


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


