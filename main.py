import serial
import serial.tools.list_ports
import tkinter as tk
from tkinter import ttk, messagebox, PhotoImage
from threading import Thread, Event
import queue
import time
import webbrowser
from pynput import keyboard
import pyautogui


class ScaleApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Приложение для работы с весами ОКБ ВЕСТА серии BM..M")

        # Настройки по умолчанию
        self.decimal_point = ','  # Запятая как десятичный разделитель
        self.after_action = 'none'  # Никаких действий после ввода
        self.hotkey = 'F2'  # Горячая клавиша по умолчанию

        # Инициализация переменных
        self.serial_port = None
        self.serial_thread = None
        self.stop_event = Event()
        self.data_queue = queue.Queue()
        self.last_weight = None
        self.last_units = None
        self.listener = None

        # Загрузка иконки GitHub
        try:
            self.github_icon = PhotoImage(file="assets/github.png").subsample(15, 15)
        except:
            self.github_icon = None

        self.create_widgets()
        self.process_data()
        self.update_ports()  # Инициализация списка портов
        self.start_hotkey_listener()

    def update_ports(self):
        """Обновляет список доступных COM-портов"""
        ports = [port.device for port in serial.tools.list_ports.comports()]
        self.port_combobox['values'] = ports
        if ports:
            self.port_combobox.current(0)

    def set_decimal_point(self):
        """Устанавливает выбранный десятичный разделитель"""
        self.decimal_point = self.decimal_var.get()
        if self.last_weight:
            self.display_weight(self.last_weight)

    def set_after_action(self):
        """Устанавливает действие после ввода данных"""
        self.after_action = self.action_var.get()

    def set_hotkey(self):
        """Устанавливает новую горячую клавишу"""
        new_hotkey = self.hotkey_combobox.get()
        if new_hotkey and new_hotkey != self.hotkey:
            self.hotkey = new_hotkey
            self.show_status(f"Горячая клавиша изменена на {self.hotkey}")
            self.start_hotkey_listener()

    def start_hotkey_listener(self):
        """Запускает обработчик горячих клавиш"""
        if self.listener:
            self.listener.stop()

        def on_press(key):
            try:
                if key == keyboard.Key.__getattribute__(keyboard.Key, self.hotkey.lower()):
                    self.input_weight()
            except AttributeError:
                pass

        self.listener = keyboard.Listener(on_press=on_press)
        self.listener.daemon = True
        self.listener.start()

    def input_weight(self):
        """Вводит текущий вес в активное окно"""
        if self.last_weight:
            try:
                # Берем значение "как есть" с весов, только меняем разделитель
                if self.decimal_point == ',':
                    weight_str = self.last_weight.replace('.', ',')
                else:
                    weight_str = self.last_weight

                # Вставляем значение в активное окно
                pyautogui.write(weight_str)

                # Дополнительное действие если нужно
                if self.after_action == 'tab':
                    pyautogui.press('tab')
                elif self.after_action == 'enter':
                    pyautogui.press('enter')

                self.show_status(f"Введено: {weight_str}" +
                                 (f", затем {self.after_action}" if self.after_action != 'none' else ""))
            except Exception as e:
                self.show_status(f"Ошибка ввода: {str(e)}", error=True)
        else:
            self.show_status("Нет данных о весе", error=True)

    def test_input(self):
        """Тестирует ввод без реального действия"""
        if self.last_weight:
            try:
                if self.decimal_point == ',':
                    weight_str = self.last_weight.replace('.', ',')
                else:
                    weight_str = self.last_weight

                action = self.after_action if self.after_action != 'none' else "без дополнительного действия"
                self.show_status(f"Тест: будет введено '{weight_str}' ({action})")
            except Exception as e:
                self.show_status(f"Ошибка: {str(e)}", error=True)
        else:
            self.show_status("Нет данных о весе", error=True)

    def toggle_connection(self):
        """Подключается или отключается от COM-порта"""
        if self.serial_port and self.serial_port.is_open:
            self.disconnect()
        else:
            self.connect()

    def connect(self):
        """Подключается к выбранному COM-порту"""
        port = self.port_combobox.get()
        if not port:
            messagebox.showerror("Ошибка", "Не выбран COM-порт")
            return

        try:
            self.serial_port = serial.Serial(
                port=port,
                baudrate=2400,
                bytesize=serial.EIGHTBITS,
                parity=serial.PARITY_EVEN,
                stopbits=serial.STOPBITS_ONE,
                timeout=1
            )

            self.stop_event.clear()
            self.serial_thread = Thread(target=self.read_serial_data)
            self.serial_thread.daemon = True
            self.serial_thread.start()

            self.connect_button.config(text="Отключиться")
            self.show_status(f"Подключено к {port}")
        except Exception as e:
            messagebox.showerror("Ошибка", f"Не удалось подключиться: {str(e)}")

    def disconnect(self):
        """Отключается от COM-порта"""
        if self.serial_port and self.serial_port.is_open:
            self.stop_event.set()
            if self.serial_thread and self.serial_thread.is_alive():
                self.serial_thread.join(timeout=1)
            self.serial_port.close()
            self.connect_button.config(text="Подключиться")
            self.show_status("Отключено от COM-порта")

    def read_serial_data(self):
        """Читает данные с COM-порта в отдельном потоке"""
        buffer = bytearray()

        while not self.stop_event.is_set():
            try:
                if self.serial_port.in_waiting > 0:
                    data = self.serial_port.read(self.serial_port.in_waiting)
                    buffer.extend(data)

                    while len(buffer) >= 29:
                        if len(buffer) >= 29 and buffer[27] == 0x0D and buffer[28] == 0x0A:
                            packet = buffer[:29]
                            buffer = buffer[29:]
                            self.data_queue.put(packet)
                        else:
                            break

                time.sleep(0.1)
            except Exception as e:
                self.data_queue.put(f"Ошибка: {str(e)}")
                break

        self.data_queue.put(None)

    def process_data(self):
        """Обрабатывает данные из очереди и обновляет интерфейс"""
        try:
            data = self.data_queue.get_nowait()

            if data is None:
                self.disconnect()
            elif isinstance(data, str) and data.startswith("Ошибка"):
                messagebox.showerror("Ошибка", data)
                self.disconnect()
            elif isinstance(data, bytearray) and len(data) == 29:
                self.process_scale_data(data)
        except queue.Empty:
            pass

        self.root.after(100, self.process_data)

    def process_scale_data(self, data):
        """Разбирает пакет данных с весов и обновляет интерфейс"""
        try:
            # Вес (байты 1-8)
            weight_str = data[0:8].decode('ascii').strip()
            self.last_weight = weight_str
            self.display_weight(weight_str)

            # Единицы измерения (байты 9-11)
            units = data[8:11].decode('ascii').strip()
            self.last_units = units if units else None
            self.unit_var.set(units if units else "---")

            # Статус (стабильный/нестабильный)
            status = "Стабильно" if units else "Нестабильно"
            self.status_var.set(status)
        except Exception as e:
            self.show_status(f"Ошибка обработки данных: {str(e)}", error=True)

    def display_weight(self, weight_str):
        """Отображает вес с учетом выбранного десятичного разделителя"""
        if self.decimal_point == ',':
            displayed_weight = weight_str.replace('.', ',')
        else:
            displayed_weight = weight_str
        self.weight_var.set(displayed_weight)

    def show_status(self, message, error=False):
        """Отображает сообщение в статусной строке"""
        if error:
            self.status_bar.config(foreground='red')
        else:
            self.status_bar.config(foreground='black')
        self.status_bar.config(text=message)
        self.root.after(5000, lambda: self.status_bar.config(
            text=f"Готов к работе. Горячая клавиша: {self.hotkey}",
            foreground='black'))

    def create_widgets(self):
        """Создает графический интерфейс"""
        # Фрейм для настроек COM-порта
        port_frame = ttk.LabelFrame(self.root, text="Настройки COM-порта")
        port_frame.grid(row=0, column=0, padx=10, pady=5, sticky="ew")

        # Выбор COM-порта
        ttk.Label(port_frame, text="Порт:").grid(row=0, column=0, padx=5, pady=2)
        self.port_combobox = ttk.Combobox(port_frame, state="readonly")
        self.port_combobox.grid(row=0, column=1, padx=5, pady=2)

        # Кнопка обновления списка портов
        ttk.Button(port_frame, text="Обновить", command=self.update_ports).grid(row=0, column=2, padx=5, pady=2)

        # Кнопки подключения/отключения
        self.connect_button = ttk.Button(port_frame, text="Подключиться", command=self.toggle_connection)
        self.connect_button.grid(row=0, column=3, padx=5, pady=2)

        # Фрейм для настроек формата
        format_frame = ttk.LabelFrame(self.root, text="Настройки формата")
        format_frame.grid(row=1, column=0, padx=10, pady=5, sticky="ew")

        # Выбор десятичного разделителя
        ttk.Label(format_frame, text="Десятичный разделитель:").grid(row=0, column=0, padx=5, pady=2)
        self.decimal_var = tk.StringVar(value=",")
        ttk.Radiobutton(format_frame, text="Запятая (,)", variable=self.decimal_var, value=",",
                        command=self.set_decimal_point).grid(row=0, column=1, padx=5, pady=2)
        ttk.Radiobutton(format_frame, text="Точка (.)", variable=self.decimal_var, value=".",
                        command=self.set_decimal_point).grid(row=0, column=2, padx=5, pady=2)

        # Выбор действия после ввода
        ttk.Label(format_frame, text="После ввода:").grid(row=0, column=3, padx=5, pady=2)
        self.action_var = tk.StringVar(value="none")
        ttk.Radiobutton(format_frame, text="Ничего", variable=self.action_var, value="none",
                        command=self.set_after_action).grid(row=0, column=4, padx=5, pady=2)
        ttk.Radiobutton(format_frame, text="Enter", variable=self.action_var, value="enter",
                        command=self.set_after_action).grid(row=0, column=5, padx=5, pady=2)
        ttk.Radiobutton(format_frame, text="Tab", variable=self.action_var, value="tab",
                        command=self.set_after_action).grid(row=0, column=6, padx=5, pady=2)

        # Фрейм для настроек горячих клавиш
        hotkey_frame = ttk.LabelFrame(self.root, text="Настройки горячих клавиш")
        hotkey_frame.grid(row=2, column=0, padx=10, pady=5, sticky="ew")

        # Выбор горячей клавиши
        ttk.Label(hotkey_frame, text="Горячая клавиша:").grid(row=0, column=0, padx=5, pady=2)
        self.hotkey_combobox = ttk.Combobox(hotkey_frame,
                                            values=['F1', 'F2', 'F3', 'F4', 'F5', 'F6',
                                                    'F7', 'F8', 'F9', 'F10', 'F11', 'F12'],
                                            state="readonly")
        self.hotkey_combobox.grid(row=0, column=1, padx=5, pady=2)
        self.hotkey_combobox.set('F2')
        ttk.Button(hotkey_frame, text="Применить", command=self.set_hotkey).grid(row=0, column=2, padx=5, pady=2)
        ttk.Button(hotkey_frame, text="Тест ввода", command=self.test_input).grid(row=0, column=3, padx=5, pady=2)

        # Фрейм для отображения данных с весов
        data_frame = ttk.LabelFrame(self.root, text="Данные с весов")
        data_frame.grid(row=3, column=0, padx=10, pady=5, sticky="nsew")

        # Поле для отображения веса
        ttk.Label(data_frame, text="Вес:").grid(row=0, column=0, padx=5, pady=2, sticky="w")
        self.weight_var = tk.StringVar(value="---")
        ttk.Label(data_frame, textvariable=self.weight_var, font=('Arial', 24)).grid(row=1, column=0, padx=5, pady=2,
                                                                                     sticky="w")

        # Поле для отображения единиц измерения
        ttk.Label(data_frame, text="Единицы измерения:").grid(row=2, column=0, padx=5, pady=2, sticky="w")
        self.unit_var = tk.StringVar(value="---")
        ttk.Label(data_frame, textvariable=self.unit_var, font=('Arial', 12)).grid(row=3, column=0, padx=5, pady=2,
                                                                                   sticky="w")

        # Статус стабильности измерения
        ttk.Label(data_frame, text="Статус:").grid(row=4, column=0, padx=5, pady=2, sticky="w")
        self.status_var = tk.StringVar(value="---")
        ttk.Label(data_frame, textvariable=self.status_var).grid(row=5, column=0, padx=5, pady=2, sticky="w")

        # Футер с кнопкой GitHub
        footer_frame = ttk.Frame(self.root)
        footer_frame.grid(row=4, column=0, sticky="se", padx=10, pady=5)

        github_btn = ttk.Button(
            footer_frame,
            image=self.github_icon,
            text=" GitHub" if self.github_icon else "GitHub",
            compound="left" if self.github_icon else None,
            command=lambda: webbrowser.open("https://github.com/Hasan175"),
            style='Toolbutton'
        )
        github_btn.pack(side="right")

        # Статусная строка
        self.status_bar = ttk.Label(self.root, text="Готов к работе. Горячая клавиша: F2", relief=tk.SUNKEN)
        self.status_bar.grid(row=5, column=0, sticky="ew", padx=10, pady=5)

        # Настройка размеров
        self.root.columnconfigure(0, weight=1)
        self.root.rowconfigure(3, weight=1)
        data_frame.columnconfigure(0, weight=1)

    def on_closing(self):
        """Обработчик закрытия окна"""
        if self.listener:
            self.listener.stop()
        self.disconnect()
        self.root.destroy()


if __name__ == "__main__":
    try:
        from pynput import keyboard
        import pyautogui
    except ImportError as e:
        print(f"Установите необходимые модули: pip install pynput pyautogui")
        exit()

    root = tk.Tk()
    app = ScaleApp(root)
    root.protocol("WM_DELETE_WINDOW", app.on_closing)
    root.mainloop()