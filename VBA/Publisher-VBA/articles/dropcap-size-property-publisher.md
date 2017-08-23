---
title: "Свойство DropCap.Size (издатель)"
keywords: vbapb10.chm5505032
f1_keywords: vbapb10.chm5505032
ms.prod: publisher
api_name: Publisher.DropCap.Size
ms.assetid: c8111c4f-7b70-76ba-5c8e-acaeb4c90be7
ms.date: 06/08/2017
ms.openlocfilehash: 0d1e0b250f09bf3af18ec25c95309b79f22b33db
ms.sourcegitcommit: 1102fd44df64f18dc0561d0b3a7103ca81e74318
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 07/26/2017
---
# <a name="dropcapsize-property-publisher"></a>Свойство DropCap.Size (издатель)

Возвращает или задает **Long** , представляющее номер строки высокой форматирование буквицы. Чтение и запись.


## <a name="syntax"></a>Синтаксис

 _выражение_. **Размер**

 переменная _expression_A, представляет собой объект- **буквицу** .


## <a name="example"></a>Пример

В этом примере форматов буквицы в диапазоне указанный текст, который является пять строк.


```vb
Sub RaisedDropCap() 
 Dim intCount As Integer 
 With ActiveDocument.Pages(1).Shapes _ 
 .AddTextbox(Orientation:=pbTextOrientationHorizontal, _ 
 Left:=100, Top:=100, Width:=100, Height:=100) _ 
 .TextFrame.TextRange 
 For intCount = 1 To 10 
 .InsertAfter NewText:="This is a test. " 
 Next intCount 
 With .DropCap 
 .Size = 5 
 .LinesUp = 2 
 End With 
 End With 
End Sub
```


