---
title: "Свойство WebListBox.ReturnDataLabel (издатель)"
keywords: vbapb10.chm4063237
f1_keywords: vbapb10.chm4063237
ms.prod: publisher
api_name: Publisher.WebListBox.ReturnDataLabel
ms.assetid: 0c9a6942-1cc7-92b6-116e-836e79560084
ms.date: 06/08/2017
ms.openlocfilehash: f6573765e70f12939a17218fedf0d424fbe2139d
ms.sourcegitcommit: 1102fd44df64f18dc0561d0b3a7103ca81e74318
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 07/26/2017
---
# <a name="weblistboxreturndatalabel-property-publisher"></a>Свойство WebListBox.ReturnDataLabel (издатель)

Возвращает или задает **строку** , представляющую текст, используемый с веб-страницы для подписи указанного веб-объект при отправке страницы. Чтение и запись.


## <a name="syntax"></a>Синтаксис

 _выражение_. **ReturnDataLabel**

 переменная _expression_A, представляет собой объект- **WebListBox** .


## <a name="example"></a>Пример

В этом примере создается новое текстовое поле Web и определяет метку для текста в текстовом поле при отправке страницы.


```vb
Sub LabelWebTextBoxControl() 
 With ActiveDocument.Pages(1).Shapes _ 
 .AddWebControl(Type:=pbWebControlSingleLineTextBox, _ 
 Left:=100, Top:=100, Width:=300, Height:=15).WebTextBox 
 .DefaultText = "Please enter your name here" 
 .Limit = 70 
 .RequiredControl = msoTrue 
 .ReturnDataLabel = "Full_Name" 
 End With 
End Sub
```


