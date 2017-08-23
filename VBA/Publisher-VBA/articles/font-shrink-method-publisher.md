---
title: "Метод Font.Shrink (издатель)"
keywords: vbapb10.chm5373991
f1_keywords: vbapb10.chm5373991
ms.prod: publisher
api_name: Publisher.Font.Shrink
ms.assetid: c5626ef2-5351-ab49-bf86-690587daed1f
ms.date: 06/08/2017
ms.openlocfilehash: ea5f3d465124c2999d2f4910f45a8c05fe57031b
ms.sourcegitcommit: 1102fd44df64f18dc0561d0b3a7103ca81e74318
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 07/26/2017
---
# <a name="fontshrink-method-publisher"></a>Метод Font.Shrink (издатель)

Уменьшает размер шрифта до следующего доступные значения. Если выделение или диапазон содержит более одного размер шрифта, размер каждого уменьшается на следующую настройку недоступны.


## <a name="syntax"></a>Синтаксис

 _выражение_. **Сжатие**

 переменная _expression_A, представляющий объект **Font** .


## <a name="remarks"></a>Заметки

Применение метода **сжатия** к тексту, уже минимальный размер, предоставляемым Microsoft Publisher (0,5 точки) не оказывает влияния.


## <a name="example"></a>Пример

В этом примере вставляет строку по мере возрастания меньшего размера Z в новый документ.


```vb
Dim shpText As Shape 
Dim trTemp As TextRange 
Dim intCount As Integer 
 
Set shpText = ActiveDocument.Pages(1).Shapes _ 
 .AddTextbox(Orientation:=pbTextOrientationHorizontal, _ 
 Left:=100, Top:=100, Width:=300, Height:=50) 
 
Set trTemp = shpText.TextFrame.TextRange 
 
With trTemp 
 .Font.Size = 45 
 .InsertAfter NewText:="ZZZZZZZZZZ" 
 For intCount = 2 To 10 
 .Characters(Start:=intCount, _ 
 Length:=11 - intCount).Font.Shrink 
 Next intCount 
End With
```


