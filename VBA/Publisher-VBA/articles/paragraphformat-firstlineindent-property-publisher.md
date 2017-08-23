---
title: "Свойство ParagraphFormat.FirstLineIndent (издатель)"
keywords: vbapb10.chm5439493
f1_keywords: vbapb10.chm5439493
ms.prod: publisher
api_name: Publisher.ParagraphFormat.FirstLineIndent
ms.assetid: 4966b30e-7629-b66d-0870-ada91c3af4f3
ms.date: 06/08/2017
ms.openlocfilehash: 0462698f674a601b8c276d01ccd5c4df5a194d16
ms.sourcegitcommit: 1102fd44df64f18dc0561d0b3a7103ca81e74318
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 07/26/2017
---
# <a name="paragraphformatfirstlineindent-property-publisher"></a>Свойство ParagraphFormat.FirstLineIndent (издатель)

Возвращает или задает **Variant** , который представляет дискового пространства (измеряется в точках) для отступа для первой строки в абзаце. Чтение и запись.


## <a name="syntax"></a>Синтаксис

 _выражение_. **FirstLineIndent**

 переменная _expression_A, представляет собой объект- **ParagraphFormat** .


### <a name="return-value"></a>Возвращаемое значение

Variant


## <a name="example"></a>Пример

В этом примере создает текстовое поле, заполняет его текстом и первой строки каждого абзаца, наполовину дюйма.


```vb
Sub IndentFirstLines() 
 Dim intCount As Integer 
 With ActiveDocument.Pages(1).Shapes _ 
 .AddTextbox(Orientation:=pbTextOrientationHorizontal, _ 
 Left:=100, Top:=100, Width:=100, Height:=100) _ 
 .TextFrame.TextRange 
 For intCount = 1 To 10 
 .InsertAfter NewText:="This is a test. " 
 Next intCount 
 .ParagraphFormat.FirstLineIndent = InchesToPoints(0.5) 
 End With 
End Sub
```


