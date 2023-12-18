Программа нужна для конвертации данных табеля сектора в word документ

Установка
клонируем репозиторий гит через - (git clone)

Настраиваем config.yaml вид конфига представлен в example.yaml - (нужно ввести туда свои учетные данные)

Создаем venv
*в терминале сделать cd в проект пример - (cd \converter_to_docx)
*затем создать окружение - (python -m venv env)
*активировать окружение - (\env\Scripts\activate.bat)

*устанавливаем зависимости - (pip install requirements --proxy http://{ЛОГИН}:{ПАРОЛЬ}@proxy.bolid.ru:3128)

Создание exe
*в терминале вводим команду - (pyinstaller --onefile --name=converter_tabel_sector_to_word main.py)

Использование exe - 
*нужно использовать экзе вместе с файлами ruzanov_gsheet.json и config.yaml. 
*так-же должны быть 2 папки 
    1 - templates с шаблоном (будет рендерится)
    2 - output это папка в которую придет конечный файл
