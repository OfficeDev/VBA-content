---
title: "Метод Font.Grow (издатель)"
keywords: vbapb10.chm5373990
f1_keywords: vbapb10.chm5373990
ms.prod: publisher
api_name: Publisher.Font.Grow
ms.assetid: 41d48db2-4a0d-6efc-80c5-c6f035e9e6ff
ms.date: 06/08/2017
ms.openlocfilehash: 83fcda6b9ba5c8f2d37aff8b2770e8f8aa5ef4f3
ms.sourcegitcommit: 1102fd44df64f18dc0561d0b3a7103ca81e74318
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 07/26/2017
---
# <a name="fontgrow-method-publisher"></a>Метод Font.Grow (издатель)

Увеличение размера шрифта до следующего доступные значения.


## <a name="syntax"></a>Синтаксис

 _выражение_. **Увеличьте размеры**

 переменная _expression_A, представляющий объект **Font** .


## <a name="remarks"></a>Заметки

Если выделение или диапазон содержит более одного размер шрифта, размер каждого увеличивается до доступен следующий параметр.


## <a name="example"></a>Пример

В этом примере увеличивается размер шрифта четвертый word в новый текстовое поле.


```vb
Sub GrowFont() 
 Dim shpText As Shape 
 Dim intResponse As Integer 
 
 Set shpText = ActiveDocument.Pages(1).Shapes.AddTextbox( _ 
 Orientation:=pbTextOrientationHorizontal, Left:=100, _ 
 Top:=100, Width:=200, Height:=100) 
 
 With shpText.TextFrame.TextRange 
 .Text = "This is a test of the Grow method." 
 Do Until intResponse = vbNo 
 intResponse = MsgBox("Do you want to increase the " &; _ 
 "size of the font?", vbYesNo) 
 If intResponse = vbYes Then 
 .Words(4).Font.Grow 
 End If 
 Loop 
 End With 
End Sub
```

В этом примере увеличивается размер шрифта выделенного текста.




```vb
Sub IncreaseFontSizeOfSelectedText() 
 If Selection.Type = pbSelectionText Then 
 Selection.TextRange.Font.Grow 
 Else 
 MsgBox "You need to select some text." 
 End If 
End Sub
```


