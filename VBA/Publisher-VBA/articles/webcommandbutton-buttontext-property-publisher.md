---
title: "Свойство WebCommandButton.ButtonText (издатель)"
keywords: vbapb10.chm3932164
f1_keywords: vbapb10.chm3932164
ms.prod: publisher
api_name: Publisher.WebCommandButton.ButtonText
ms.assetid: 0a9a7bd9-de7e-7e80-0aa2-7cefda17f354
ms.date: 06/08/2017
ms.openlocfilehash: 7694b52ff2d1a0ca27660c9c2eb992c84a186bb7
ms.sourcegitcommit: 1102fd44df64f18dc0561d0b3a7103ca81e74318
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 07/26/2017
---
# <a name="webcommandbuttonbuttontext-property-publisher"></a>Свойство WebCommandButton.ButtonText (издатель)

Возвращает или задает **строку** , представляющую текст, отображаемый на кнопке Web. Чтение и запись.


## <a name="syntax"></a>Синтаксис

 _выражение_. **ButtonText**

 переменная _expression_A, представляет собой объект- **WebCommandButton** .


### <a name="return-value"></a>Возвращаемое значение

String


## <a name="example"></a>Пример

В этом примере создается новая кнопка команды Web, назначает текст для отображения на поверхность и адрес электронной почты, к которому следует отправить данные формы.


```vb
Sub NewWebForm() 
 With ActiveDocument.Pages.Add(Count:=1, After:=1) 
 With .Shapes.AddWebControl(Type:=pbWebControlCommandButton, _ 
 Left:=72, Top:=72, Width:=72, Height:=36) 
 With .WebCommandButton 
 .ButtonType = pbCommandButtonSubmit 
 .ButtonText = "Send Form:" 
 .EmailAddress = "someone@example.com" 
 End With 
 End With 
 End With 
End Sub
```


