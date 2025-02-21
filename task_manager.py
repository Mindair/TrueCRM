# task_manager.py

from datetime import datetime
from openpyxl import Workbook, load_workbook
from openpyxl.utils.exceptions import InvalidFileException
from openpyxl.styles import PatternFill
import json


class TaskManager:
    def __init__(self, filename="tasks.xlsx"):
        self.tasks = []  # Инициализируем пустой список задач
        self.filename = filename  # Сохраняем имя файла

        # Попытка загрузить задачи из существующего файла
        try:
            wb = load_workbook(self.filename)
            ws = wb.active

            # Пропускаем заголовки и читаем данные
            for row in ws.iter_rows(min_row=2, values_only=True):  # Начинаем с второй строки
                if row[0] is not None:  # Проверяем, что номер задачи указан
                    task = {
                        "description": row[3],  # Описание (четвертый столбец)
                        "done": True if row[1] == "Выполнено" else False,  # Статус (второй столбец)
                        "company": row[2],  # Компания (третий столбец)
                        "deadline": row[4] if row[4] != "Нет дедлайна" else None,  # Дедлайн (пятый столбец)
                        "created_at": row[5] if row[5] else None  # Время создания (шестой столбец)
                    }
                    self.tasks.append(task)

        except FileNotFoundError:
            print("Файл задач не найден. Создан новый список задач.")
        except InvalidFileException:
            print("Ошибка чтения файла. Создан новый список задач.")

    def addTask(self):
        description = input("Введите описание задачи: ")
        company = input("Введите название компании: ")
        while True:
            deadline_input = input("Введите дедлайн (YYYY-MM-DD) или нажмите Enter, если нет дедлайна: ").strip()
            if not deadline_input:
                deadline = None
                print("Дедлайн не установлен.")
                break
            else:
                try:
                    if len(deadline_input) != 10 or deadline_input[4] != '-' or deadline_input[7] != '-':
                        raise ValueError("Неверный формат даты. Используйте формат YYYY-MM-DD.")
                    deadline_date = datetime.strptime(deadline_input, '%Y-%m-%d')
                    deadline = deadline_date.strftime('%Y-%m-%d')
                    current_date = datetime.now().date()
                    if deadline_date.date() < current_date:
                        print("Дедлайн должен быть в будущем.")
                        continue
                    print(f"Дедлайн успешно установлен: {deadline}")
                    break
                except ValueError as e:
                    print(f"Ошибка: {e}. Пожалуйста, используйте формат YYYY-MM-DD.")

        # Автоматическое добавление времени создания
        created_at = datetime.now().strftime('%Y-%m-%d %H:%M:%S')  # Текущая дата и время

        task = {
            "description": description,
            "done": False,
            "company": company,
            "deadline": deadline,
            "created_at": created_at  # Добавляем время создания
        }
        self.tasks.append(task)
        print("Задача добавлена!")

    def lookTasks(self):
        if not self.tasks:
            print("Список задач пуст.")
            return

        print("\nВсе задачи:")
        for i, task in enumerate(self.tasks, start=1):
            status = "[v]" if task["done"] else "[x]"
            deadline = f"(Дедлайн: {task['deadline']})" if task["deadline"] else ""
            created_at = f"(Создана: {task['created_at']})"  # Выводим время создания
            print(f"{i}. {status} Компания: {task['company']} | Задача: {task['description']} {deadline} {created_at}")

    def removeTask(self):
        self.lookTasks()
        if not self.tasks:
            return
        try:
            number = int(input("Введите номер задачи для удаления: "))
            if 1 <= number <= len(self.tasks):
                removed_task = self.tasks.pop(number - 1)
                print(f"Задача \"{removed_task['description']}\" удалена.")
            else:
                print("Неверный номер задачи.")
        except ValueError:
            print("Введите корректный номер.")

    def markTask(self):
        self.lookTasks()
        if not self.tasks:
            return
        try:
            number = int(input("Введите номер задачи для отметки как выполненной: "))
            if 1 <= number <= len(self.tasks):
                self.tasks[number - 1]["done"] = True
                print("Задача отмечена как выполненная.")
            else:
                print("Неверный номер задачи.")
        except ValueError:
            print("Введите корректный номер.")

    def clearTasks(self):
        confirmation = input("Вы уверены, что хотите удалить ВСЕ задачи? (да/нет): ").strip().lower()
        if confirmation == "да":
            self.tasks = []  # Очищаем список задач
            print("Все задачи успешно удалены.")

            # Сохраняем пустой список в Excel
            try:
                wb = Workbook()
                ws = wb.active
                ws.title = "Задачи"

                # Обновленные заголовки
                headers = ["Номер", "Статус", "Компания", "Описание", "Дедлайн", "Время создания"]
                ws.append(headers)

                # Сохраняем пустой файл
                wb.save(self.filename)
                print(f"Файл {self.filename} очищен.")
            except Exception as e:
                print(f"Ошибка при очистке файла: {e}")
        else:
            print("Операция отменена.")

    def save_to_json(self, filename="tasks.json"):
        with open(filename, "w", encoding="utf-8") as file:
            json.dump(self.tasks, file, indent=4, ensure_ascii=False)
        print(f"Задачи успешно сохранены в {filename}.")

    def save_to_txt(self, filename="tasks.txt"):
        with open(filename, "w", encoding="utf-8") as file:
            if not self.tasks:
                file.write("Список задач пуст.")
            else:
                for i, task in enumerate(self.tasks, start=1):
                    status = "[v]" if task["done"] else "[x]"
                    deadline = f"(Дедлайн: {task['deadline']})" if task["deadline"] else ""
                    created_at = f"(Создана: {task['created_at']})"  # Выводим время создания
                    file.write(
                        f"{i}. {status} Компания: {task['company']} | Задача: {task['description']} {deadline} {created_at}\n")
        print(f"Задачи успешно сохранены в {filename}.")

    def save_to_excel(self, filename="tasks.xlsx"):
        try:
            # Попытка загрузить существующий файл
            wb = load_workbook(filename)
            ws = wb.active
        except FileNotFoundError:
            # Если файл не найден, создаем новый
            wb = Workbook()
            ws = wb.active
            ws.title = "Задачи"

            # Обновленные заголовки
            headers = ["Номер", "Статус", "Компания", "Описание", "Дедлайн", "Время создания"]
            ws.append(headers)

        # Цветовая схема
        green_fill = PatternFill(start_color="C6EFCE", end_color="C6EFCE",
                                 fill_type="solid")  # Зеленый для выполненых задач
        red_fill = PatternFill(start_color="FFC7CE", end_color="FFC7CE",
                               fill_type="solid")  # Красный для невыполненых задач

        # Определяем номер следующей задачи
        if ws.max_row > 1:  # Если есть уже записанные задачи
            last_number = ws.cell(row=ws.max_row, column=1).value  # Получаем последний номер
        else:
            last_number = 0  # Если файл пустой, начинаем с 1

        # Добавляем новые задачи
        for task in self.tasks:
            last_number += 1
            status = "Выполнено" if task["done"] else "Не выполнено"
            deadline = task["deadline"] if task["deadline"] else "Нет дедлайна"
            row = [
                last_number,
                status,
                task["company"],
                task["description"],
                deadline,
                task["created_at"]  # Добавляем время создания
            ]
            ws.append(row)

            # Применяем цветовую заливку только к ячейке "Статус"
            row_index = ws.max_row  # Номер текущей строки
            status_cell = ws.cell(row=row_index, column=2)  # Ячейка "Статус" (второй столбец)
            if task["done"]:
                status_cell.fill = green_fill  # Зеленый для выполненых задач
            else:
                status_cell.fill = red_fill  # Красный для невыполненых задач

        # Сохраняем файл
        wb.save(filename)
        print(f"Задачи успешно сохранены в {filename}.")