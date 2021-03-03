Python модуль по работе с документами с расширением `docx`. 
Используется для вставки данных в шаблон. 
***
Добавление полей в шаблон:

Для этого в программе MS Office во вкладке "Вставка" в меню "Экспресс блоки" выбрать "поле". 
Откроется меню где нужно выбрать тип поля и ввести имя поля. В шаблоне, поля, которые должны заполняться, должы иметь модификацию `Mergefield`.
Имя произвольное
***
Для вставки данных в шаблон:

Для вставки данных в шаблон нужно сгенерировать словарь, в котором:
ключь - это название поля в шаблоне
значение - это данные которые нужно вставить в это поле
***
Использование 

на вход подается подготовленный шаблон `template` и словарь `inserted_data`.
`a = Docx_insert_data_class.docx_insert_data_class(template, inserted_data)`

при запуске проекта генерируется web-страничка в которой кнопка. 
При нажатии на кнопку сгенерированный файл с вставленными данными сохраняется на ваш ПК




