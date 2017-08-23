---
title: "Свойство Window.Caption (издатель)"
keywords: vbapb10.chm262146
f1_keywords: vbapb10.chm262146
ms.prod: publisher
api_name: Publisher.Window.Caption
ms.assetid: 1dbf66c9-e964-b17f-684f-70cbbaa5fbc7
ms.date: 06/08/2017
ms.openlocfilehash: 217508325c9d66d2f35475f103e6733e1acb307e
ms.sourcegitcommit: 1102fd44df64f18dc0561d0b3a7103ca81e74318
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 07/26/2017
---
# <a name="windowcaption-property-publisher"></a>Свойство Window.Caption (издатель)

Возвращает или задает **строку** , определяющую заголовок в верхней части окна приложения Microsoft Publisher. Чтение и запись.


## <a name="syntax"></a>Синтаксис

 _выражение_. **Заголовок**

 переменная _expression_A, представляющий объект **Window** .


### <a name="return-value"></a>Возвращаемое значение

String


## <a name="example"></a>Пример

В следующем примере показано, как подпрограмма может временно изменение заголовка окна Publisher, а затем восстановить его позже.


```vb
Sub WindowCaption() 
 Dim strCaption As String 
 
 strCaption = ActiveWindow.Caption 
 
 ActiveWindow.Caption = "Custom process--please wait..." 
 
 ' Run custom code here. 
 
 ActiveWindow.Caption = strCaption 
End Sub
```


