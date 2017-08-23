---
title: "Метод TextFrame.BreakForwardLink (издатель)"
keywords: vbapb10.chm3866661
f1_keywords: vbapb10.chm3866661
ms.prod: publisher
api_name: Publisher.TextFrame.BreakForwardLink
ms.assetid: 60a7a798-ebd3-e00d-032d-685dd0d5a042
ms.date: 06/08/2017
ms.openlocfilehash: 438f78b71c4fd60e05508d24aed341fad6895253
ms.sourcegitcommit: 1102fd44df64f18dc0561d0b3a7103ca81e74318
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 07/26/2017
---
# <a name="textframebreakforwardlink-method-publisher"></a>Метод TextFrame.BreakForwardLink (издатель)

Прерывается прямая ссылка для frame указанный текст, если существует таких ссылок.


## <a name="syntax"></a>Синтаксис

 _выражение_. **BreakForwardLink**

 переменная _expression_A, представляет собой объект- **TextFrame** .


## <a name="remarks"></a>Заметки

Этот метод для применения к фигуры середину цепочки фигур с помощью связанных текстовых кадров разрываются цепочки, отправляемых из двух наборов связанные фигуры. Весь текст, тем не менее, останутся первого ряда связанные фигуры.


## <a name="example"></a>Пример

В этом примере создается новая публикация, добавляет цепочки три связанных текстовых полей и затем нарушает связь после вторым текстовым полем.


```vb
Sub BreakTextLink() 
 Dim shpTextbox1 As Shape 
 Dim shpTextbox2 As Shape 
 Dim shpTextbox3 As Shape 
 
 Set shpTextbox1 = ActiveDocument.Pages(1).Shapes.AddTextbox _ 
 (Orientation:=msoTextOrientationHorizontal, _ 
 Left:=72, Top:=36, Width:=72, Height:=36) 
 shpTextbox1.TextFrame.TextRange = "This is some text. " _ 
 &; "This is some more text. This is even more text. " _ 
 &; "And this is some more text and even more text." 
 
 Set shpTextbox2 = ActiveDocument.Pages(1).Shapes.AddTextbox _ 
 (Orientation:=msoTextOrientationHorizontal, _ 
 Left:=72, Top:=108, Width:=72, Height:=36) 
 
 Set shpTextbox3 = ActiveDocument.Pages(1).Shapes.AddTextbox _ 
 (Orientation:=msoTextOrientationHorizontal, _ 
 Left:=72, Top:=180, Width:=72, Height:=36) 
 
 shpTextbox1.TextFrame.NextLinkedTextFrame = shpTextbox2.TextFrame 
 shpTextbox2.TextFrame.NextLinkedTextFrame = shpTextbox3.TextFrame 
 MsgBox "Textboxes 1, 2, and 3 are linked." 
 shpTextbox2.TextFrame.BreakForwardLink 
End Sub
```


