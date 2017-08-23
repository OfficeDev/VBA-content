---
title: "Свойство WebOptionButton.ReturnDataLabel (издатель)"
keywords: vbapb10.chm4259843
f1_keywords: vbapb10.chm4259843
ms.prod: publisher
api_name: Publisher.WebOptionButton.ReturnDataLabel
ms.assetid: 22b4a4d6-1068-2b35-d054-42bbea3f9098
ms.date: 06/08/2017
ms.openlocfilehash: f721f98134a1959a2661e2c16d7d457b20b3afe4
ms.sourcegitcommit: 1102fd44df64f18dc0561d0b3a7103ca81e74318
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 07/26/2017
---
# <a name="weboptionbuttonreturndatalabel-property-publisher"></a>Свойство WebOptionButton.ReturnDataLabel (издатель)

Возвращает или задает **строку** , представляющую текст, используемый с веб-страницы для подписи указанного веб-объект при отправке страницы. Чтение и запись.


## <a name="syntax"></a>Синтаксис

 _выражение_. **ReturnDataLabel**

 переменная _expression_A, представляет собой объект- **WebOptionButton** .


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


