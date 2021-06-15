'Тестовые данные
Dim name,page,pagename,loadfile
name = "T-AXR_1011"
page = "Лимиты Agile"
pagename = ".*Бюджет трайба.*"
loadfile = "C:\TEMP\ROTBS.xlsx" 

'1. Авторизация
Setting.WebPackage("ReplayType")=1
LOGINSUDIR (name)
wait(3)

'2. Проверка начальной страницы
set GH = Browser("name:=Начальная страница").Page("title:=Начальная страница")
i = 0
Do while not GH.Exist = true
i = i + 1
If i = 10 Then
     Reporter.ReportEvent micFail, "Начальная страница", "Начальная страница не отображается - Пользователь *1011"
     ExitTest
End  If
Loop 
'Проверка пользователя
a = GH.WebButton("acc_name:=Моя область").GetROProperty("title")
If a = "Тест Тест22019" then
    Reporter.ReportEvent micPass, "Пользователь корректный", "Выполнено"
else
    Reporter.ReportEvent micFail, "Не корректный пользователь", "Не выполнено"
    ExitTest
End If
wait(3)
'3. Открыть плитку
Set GHN = Browser("name:=Начальная страница").Page("title:=Начальная страница")
GHN.WebList("acc_name:=Групповая навигация","first item:=Моя начальная страница").WebElement("html tag:=DIV","innerhtml:=Лимиты Agile","innertext:=Лимиты Agile").Click
GHN.WebButton("acc_name:=.*Годовое планирование бюджета Agile.*","name:=Годовое планирование бюджета Agile").Click

'4. Окно запросы и переход в рабочую область
zaprosyAGILE page,pagename
wait(3)

'5. Загрузка шаблона
ShablonLoad loadfile,page
  
'6. Заполнение "Общий фильтр"
'Трайб
tr = DataTable.Value("TR", dtGlobalSheet)
obshfiltr page, "__input2-vhi", "Трайб", tr
wait(3)
'Команда
tm = DataTable.Value("TM", dtGlobalSheet)
obshfiltr page, "__input3-vhi", "Команда", tm
wait(3)
'Продукт
pr = DataTable.Value("PR", dtGlobalSheet)
obshfiltr page, "__input6-vhi", ".*Продукт.*", pr
wait(3)

'7. Переход в монитор согласования
Set AG = Browser("name:="&page&"").Page("title:="&page&"").Frame("html id:=openDocChildFrame")
AG.WebElement("html id:=ICON_TOGGLE_MONITOR_DISPLAY_control").Click
i = 0
Do while AG.WebElement("innerhtml:=Монитор согласования","innertext:=Монитор согласования","outertext:=Монитор согласования").exist = False and i < 10
i = i + 1
Loop

'8. Фильтрация по БА
BE = "9900"
tipr = DataTable.Value("TIPR", dtGlobalSheet)
BA = ""&BE&tr&tm&"1000РОТ"&pr&tipr&""
BAI = "innertext:="&BE&tr&tm&"1000РОТ"&pr&tipr&""
BAO = "outertext:="&BE&tr&tm&"1000РОТ"&pr&tipr&""
Ozidanie(page)
BAFILTR BA,page,BAI,BAO
Ozidanie(page)
'9. Проверка статуса
statuspoz "Лимиты Agile", "Отклонен"

'10. Отправить на согласование
AG.WebElement("html id:=ICON_WF_SEND_ENABLED_control").Click
i = 0
Do while not AG.WebElement("innerhtml:=Комментарии").exist = True 
i = i + 1
If i = 10 Then
     Reporter.ReportEvent micFail, "Окно комменарии", "Окно комментарии не отображается"
     ExitTest
End  If
Loop
AG.SAPUIButton("name:=ОК").Click
'Проверка статуса отправки
i = 0
Do while AG.WebTable("html id:=WF_MONITOR_table").WebTable("html id:=WF_MONITOR_rowHeaderArea_container").WebTable("html id:=WF_MONITOR_rowHeaderArea").WebElement("html tag:=TD","innertext:=Отправлен на согласование","outertext:=Отправлен на согласование").exist = false and i < 5
i = i + 1
wait (5)
Loop
If i < 5 Then
	Reporter.ReportEvent micPass, "Отправлено на согласование", "Успешно"
Else
    Reporter.ReportEvent micFail, "Не отправлено на согласование", "Ошибка"
End If

'11. Проверка статуса
statuspoz "Лимиты Agile", "Отправлен на согласование"

'Закрыть Chrome
SystemUtil.CloseProcessByName("chrome.exe")
