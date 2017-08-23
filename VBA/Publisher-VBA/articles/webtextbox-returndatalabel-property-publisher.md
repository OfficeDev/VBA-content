---
title: "Свойство WebTextBox.ReturnDataLabel (издатель)"
keywords: vbapb10.chm4194311
f1_keywords: vbapb10.chm4194311
ms.prod: publisher
api_name: Publisher.WebTextBox.ReturnDataLabel
ms.assetid: 83beba69-3d04-2010-0656-d6a27c08951c
ms.date: 06/08/2017
ms.openlocfilehash: 29b57ba230a4d121b7caee4a4b152c18a829f35e
ms.sourcegitcommit: 1102fd44df64f18dc0561d0b3a7103ca81e74318
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 07/26/2017
---
# <a name="webtextboxreturndatalabel-property-publisher"></a>Свойство WebTextBox.ReturnDataLabel (издатель)

Возвращает или задает **строку** , представляющую текст, используемый с веб-страницы для подписи указанного веб-объект при отправке страницы. Чтение и запись.


## <a name="syntax"></a>Синтаксис

 _выражение_. **ReturnDataLabel**

 переменная _expression_A, представляет собой объект- **WebTextBox** .


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


