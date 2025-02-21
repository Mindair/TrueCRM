# main.py

from task_manager import TaskManager


def main():
    manager = TaskManager()

    while True:
        print("\nМеню:")
        print("1. Добавить задачу")
        print("2. Просмотреть задачи")
        print("3. Отметить задачу как выполненную")
        print("4. Удалить задачу")
        print("5. Выйти")
        print("6. Сохранить задачи")
        print("7. Отфильтровать задачи")
        print("8. Очистить все задачи")

        choice = input("Выберите действие: ")

        if choice == "1":
            manager.addTask()
        elif choice == "2":
            manager.lookTasks()
        elif choice == "3":
            manager.markTask()
        elif choice == "4":
            manager.removeTask()
        elif choice == "5":
            print("Выход из программы.")
            break
        elif choice == "6":
            save_choice = input("Выберите формат сохранения (json/txt/excel): ").lower()
            if save_choice == "json":
                manager.save_to_json()
            elif save_choice == "txt":
                manager.save_to_txt()
            elif save_choice == "excel":
                manager.save_to_excel()
            else:
                print("Неверный формат.")
        elif choice == "7":
            manager.filter_by_company()
        elif choice == "8":
            manager.clearTasks()
        else:
            print("Неверный выбор. Попробуйте снова.")


if __name__ == "__main__":
    main()