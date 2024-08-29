import os
from tkinter import filedialog

import customtkinter
import pystray
import pywinstyles
from CTkMenuBar import CTkTitleMenu, CustomDropdownMenu
from CTkMessagebox import CTkMessagebox
from PIL import Image
from pystray import Menu, MenuItem

from db import Server, connection
from utils import read_excel, resource, write_excel

# customtkinter.set_appearance_mode(
#     "System"
# )  # Modes: "System" (standard), "Dark", "Light"
# customtkinter.set_default_color_theme(
#     "dark-blue"
# )  # Themes: "blue" (standard), "green", "dark-blue"


__version__ = "2.3"

WIDTH = 530
HEIGHT = 230


class ServerManagerGUI(customtkinter.CTk):
    def __init__(self):
        super().__init__()

        self.title("")
        # 使得 tabview 填充整个窗口
        self.grid_rowconfigure(0, weight=1)  # configure grid system
        self.grid_columnconfigure(0, weight=1)

        x, y = (
            (self.winfo_screenwidth() - WIDTH) // 2,
            (self.winfo_screenheight() - HEIGHT) // 2,
        )

        self.geometry(f"{WIDTH}x{HEIGHT}+{x}+{y}")
        self.resizable(0, 0)
        self.protocol("WM_DELETE_WINDOW", self.on_exit)

        self.font = customtkinter.CTkFont(family="楷体", size=16, weight="normal")

        self.tabview = customtkinter.CTkTabview(master=self)
        self.tabview.grid(row=0, column=0, padx=(15, 15), pady=(0, 10), sticky="NSEW")

        self.create_menu()
        self.create_server_list_tab()
        self.create_server_path_tab()

        self.show()

    # 创建 menu
    def create_menu(self):
        menu = CTkTitleMenu(
            master=self, title_bar_color="transparent", padx=0, x_offset=60
        )
        file_menu = menu.add_cascade("文件", fg_color="transparent")
        theme_menu = menu.add_cascade("主题", fg_color="transparent")
        help_menu = menu.add_cascade("帮助", fg_color="transparent")

        file_dropdown = CustomDropdownMenu(file_menu)
        file_dropdown.add_option("导入", command=self.import_excel)
        file_dropdown.add_option("导出", command=self.export_excel)

        help_dropdown = CustomDropdownMenu(help_menu)
        help_dropdown.add_option("文档", command=self.open_doc)

        theme_dropdown = CustomDropdownMenu(theme_menu)
        theme_dropdown.add_option("默认", command=self.change_theme("normal"))
        theme_dropdown.add_option("optimised", command=self.change_theme("optimised"))
        theme_dropdown.add_option("win7", command=self.change_theme("win7"))
        theme_dropdown.add_option("dark", command=self.change_theme("dark"))
        # theme_dropdown.add_option("inverse", command=self.change_theme("inverse"))
        # theme_dropdown.add_option("mica", command=self.change_theme("mica"))

    def create_server_list_tab(self):
        server_list_tab = self.tabview.add(name="服务器列表")

        project = customtkinter.CTkLabel(
            server_list_tab, text="项目名称", font=self.font
        )
        project.grid(row=1, column=0, padx=(8, 10), pady=(10, 5), sticky="W")

        self.comboBox = customtkinter.CTkComboBox(
            server_list_tab,
            values=[],
            width=380,
            state="readonly",
            font=self.font,
            command=self.refresh_server_list,
        )
        self.comboBox.grid(row=1, column=1, padx=(8, 10), pady=(10, 5), sticky="W")

        data_class = customtkinter.CTkLabel(
            server_list_tab, text="数据类型", font=self.font
        )
        data_class.grid(row=2, column=0, padx=(8, 10), pady=(10, 5), sticky="W")

        self.data_class_comboBox = customtkinter.CTkComboBox(
            server_list_tab,
            values=[],
            width=380,
            state="readonly",
            font=self.font,
            command=self.refresh_path_list,
        )
        self.data_class_comboBox.grid(
            row=2, column=1, padx=(8, 10), pady=(10, 5), sticky="W"
        )

        path = customtkinter.CTkLabel(server_list_tab, text="数据地址", font=self.font)
        path.grid(row=3, column=0, padx=(8, 10), pady=(5, 5), sticky="W")

        self.path_comboBox = customtkinter.CTkComboBox(
            server_list_tab,
            values=[],
            state="readonly",
            font=self.font,
            width=380,
        )
        self.path_comboBox.grid(row=3, column=1, padx=(8, 10), pady=(5, 5), sticky="W")

        delete = customtkinter.CTkButton(
            server_list_tab,
            text="删除",
            font=self.font,
            width=10,
            command=self.delete_server,
        )
        delete.grid(
            row=4, column=0, columnspan=2, padx=(10, 70), pady=(8, 10), sticky="E"
        )

        open = customtkinter.CTkButton(
            server_list_tab,
            text="打开",
            font=self.font,
            width=10,
            command=self.open_server,
        )
        open.grid(row=4, column=1, padx=(10, 10), pady=(8, 10), sticky="E")

    def create_server_path_tab(self):
        server_path_tab = self.tabview.add(name="添加服务器")

        project = customtkinter.CTkLabel(
            server_path_tab, text="项目名称", font=self.font
        )
        project.grid(row=1, column=0, padx=(8, 10), pady=(10, 5), sticky="W")

        self.comboBox2 = customtkinter.CTkComboBox(
            server_path_tab,
            values=[],
            width=380,
            font=self.font,
            command=self.refresh_server_list,
        )
        self.comboBox2.grid(
            row=1, column=1, columnspan=2, padx=(8, 10), pady=(10, 5), sticky="W"
        )
        self.comboBox2.set("")

        data_class = customtkinter.CTkLabel(
            server_path_tab, text="数据类型", font=self.font
        )
        data_class.grid(row=2, column=0, padx=(8, 10), pady=(10, 5), sticky="W")

        self.data_class_comboBox2 = customtkinter.CTkComboBox(
            server_path_tab,
            values=[],
            width=380,
            font=self.font,
        )
        self.data_class_comboBox2.grid(
            row=2, column=1, columnspan=2, padx=(8, 10), pady=(10, 5), sticky="W"
        )
        self.data_class_comboBox2.set("")

        path = customtkinter.CTkLabel(server_path_tab, text="数据地址", font=self.font)
        path.grid(row=3, column=0, padx=(8, 10), pady=(5, 5), sticky="W")

        self.entry = customtkinter.CTkEntry(
            server_path_tab,
            font=self.font,
            width=325,
        )
        self.entry.grid(row=3, column=1, padx=(8, 10), pady=(5, 5), sticky="W")
        self.choose_button = customtkinter.CTkButton(
            server_path_tab,
            text="选择",
            width=10,
            font=self.font,
            command=self.choose_folder,
        )
        self.choose_button.grid(row=3, column=2, padx=(3, 5), pady=(5, 5), sticky="W")

        self.add = customtkinter.CTkButton(
            server_path_tab,
            text="添加",
            font=self.font,
            width=10,
            command=self.add_server,
        )
        self.add.grid(
            row=4, column=0, columnspan=3, padx=(5, 5), pady=(8, 10), sticky="E"
        )

    # 展示数据
    def show(self):
        with connection() as conn:
            names = conn.select_names()

            # 如果没有数据，则清空 combobox 和 entry
            if not names:
                self.comboBox.configure(values=[])
                self.data_class_comboBox.configure(values=[])
                self.path_comboBox.configure(values=[])
                self.comboBox.set(value="")
                self.data_class_comboBox.set(value="")
                self.path_comboBox.set(value="")

                self.comboBox2.configure(values=[])
                self.comboBox2.set(value="")
                self.data_class_comboBox2.configure(values=[])
                self.data_class_comboBox2.set("")
                self.entry.delete(0, "end")

                return

            self.comboBox.configure(values=names)
            self.comboBox2.configure(values=names)
            # 查询 last_selected 为 1 的记录
            server = conn.select_by_last_selected()

            if not server:
                name = names[0]
                server = conn.select_by_name(name)

            name, data_class, path = server.name, server.data_class, server.path
            self.comboBox.set(name)

            self.data_class_comboBox.set(value=data_class)
            self.path_comboBox.set(path)

            data_classes = conn.select_data_classes_by_name(name)
            paths = conn.select_paths(name, data_class)

            self.data_class_comboBox.configure(values=data_classes)
            self.data_class_comboBox2.configure(values=data_classes)
            self.path_comboBox.configure(values=paths)

    # 切换数据类型时刷新路径列表
    def refresh_path_list(self, data_class: str) -> None:
        name = self.comboBox.get()
        with connection() as conn:
            paths = conn.select_paths(name, data_class)

        self.path_comboBox.configure(values=paths)
        self.path_comboBox.set(paths.pop())

    # 切换项目名称时刷新数据列表
    def refresh_server_list(self, name):
        with connection() as conn:
            data_classes = conn.select_data_classes_by_name(name)

            data_class = data_classes[-1]
            if self.tabview.get() == "添加服务器":
                self.data_class_comboBox2.configure(values=data_classes)
                self.data_class_comboBox2.set(data_class)

                return

            paths = conn.select_paths(name, data_class)

            self.data_class_comboBox.configure(values=data_classes)
            self.data_class_comboBox.set(data_class)
            self.path_comboBox.configure(values=paths)
            self.path_comboBox.set(paths[0])

    # 打开服务器
    def open_server(self):
        path = self.path_comboBox.get()
        os.startfile(path)

        with connection() as conn:
            conn.update_last_selected()
            conn.update_last_selected_by_path(path)

    # 删除服务器
    def delete_server(self):
        path = self.path_comboBox.get()

        with connection() as conn:
            conn.delete_by_path(path)

        ServerManagerGUI.show_info("Successful", "删除成功!", "check")
        self.show()

    # 添加服务器
    def add_server(self):
        name = self.comboBox2.get()
        data_class = self.data_class_comboBox2.get()
        path = self.entry.get()

        with connection() as conn:
            if conn.select_id_by_path(path):
                ServerManagerGUI.show_info("Error", "该路径已添加！", "warning")

                return

            server = Server(name, data_class, path)
            conn.insert(server)

        ServerManagerGUI.show_info("Successful", "添加成功!", "check")
        self.show()

    def choose_folder(self) -> None:
        self.entry.delete(0, "end")

        path = filedialog.askdirectory(title="选择文件夹", parent=self)
        self.entry.insert(0, path)

    def open_doc(self):
        import webbrowser

        doc = resource("./static/doc.pdf")
        webbrowser.open(doc)

    def export_excel(self) -> None:
        path = filedialog.askdirectory(title="选择文件夹", parent=self)

        if not path:
            return

        sql = """
            SELECT id, name, data_class, path, createAt
            FROM servers
        """

        with connection() as conn:
            data = list(conn.execute(sql))

        try:
            write_excel(data, path)

        except Exception as e:
            ServerManagerGUI.show_info("Error", f"导出失败: {e}", "cancel")
            return

        ServerManagerGUI.show_info("Successful", "导出成功!", "check")

    def import_excel(self) -> None:
        path = filedialog.askopenfilename(
            title="选择文件", filetypes=[("EXCEL Files", "*.xlsx *.xls")], parent=self
        )

        if not path:
            return

        data = read_excel(path)
        try:
            with connection() as conn:
                for _, name, data_class, path, _ in data:
                    if not conn.select_id_by_path(path):
                        server = Server(name, data_class, path)
                        conn.insert(server)

        except Exception as e:
            ServerManagerGUI.show_info("Error", f"导入失败: {e}", "cancel")
            return

        ServerManagerGUI.show_info("Successful", "导入成功!", "check")
        self.show()

    def change_theme(self, theme: str) -> None:
        def inner():
            if theme != "normal":
                pywinstyles.apply_style(self, "normal")

            pywinstyles.apply_style(self, theme)

        return inner

    @staticmethod
    def show_info(title: str, message: str, icon: str) -> None:
        CTkMessagebox(
            title=title,
            message=message,
            icon=icon,
            option_1="确定",
            button_width=50,
        )

    def quit_window(self, icon: pystray._base.Icon):
        icon.stop()
        self.destroy()

    def show_window(self):
        self.deiconify()

    def on_exit(self):
        self.withdraw()


if __name__ == "__main__":
    app = ServerManagerGUI()

    menu = (
        MenuItem("显示", app.show_window, default=True),
        Menu.SEPARATOR,
        MenuItem("退出", app.quit_window),
    )
    image = Image.open(resource("./static/logo.ico"))
    icon = pystray.Icon("icon", image, "服务器管理工具", menu)

    from threading import Thread

    Thread(target=icon.run, daemon=True).start()
    app.iconbitmap(default=resource("./static/logo.ico"))
    app.mainloop()
