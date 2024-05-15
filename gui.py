import os
import re
import time
import webbrowser
import dns.resolver
from base64 import b64decode
from datetime import datetime
from typing_extensions import Literal

import tkinter as tk
from tkinter import ttk
from tkinter import filedialog
from tkinter import messagebox

import utilities
from utilities import IPTranslator
from utilities import PropagatingThread
from utilities import VerticalScrolledFrame


def browse_file(string_var: tk.StringVar,
                files_ext: list | tuple | set,
                title: str = "Select a file") -> str | None:
    files_ext: str = ";".join([f"*.{x}" for x in set(files_ext)])

    # Open a file dialog
    filename = filedialog.askopenfilename(title=title,
                                          filetypes=(("Supported Files", files_ext),
                                                     ("All Files", "*.*")))

    # Check if a file is selected
    if filename != "":
        # Insert the file path in the entry
        string_var.set(filename)
        return filename


def separator(parent: tk.Tk | tk.Toplevel | ttk.Frame,
              orient: Literal["horizontal", "vertical"] = "horizontal",
              fill: Literal["x", "y", "both", "none"] = "x",
              y_offset: str | float | tuple[str | float, str | float] = (0, 10),
              x_offset: str | float | tuple[str | float, str | float] = 20) -> None:
    """Create a separator widget."""
    __separator = ttk.Separator(parent, orient=orient)
    __separator.pack(fill=fill, pady=y_offset, padx=x_offset)


class EntryRow:
    def __init__(self, parent: ttk.Frame, row: int, **kwargs):
        """Create a row for each method in the application."""

        # Create  a Label
        self.label = ttk.Label(parent, text=kwargs["label"]["text"])
        self.label.grid(row=row, column=0, sticky="e", padx=(0, 5), pady=(0, 5))

        def entry_callback(*args) -> None:
            kwargs["entry"]["callback"](str(self.entry.focus_get()))

        self.string_var = tk.StringVar()
        if kwargs["entry"].get("callback"):
            self.string_var.trace_add("write", entry_callback)

        # Create an Entry
        self.entry = ttk.Entry(parent,
                               width=kwargs["entry"]["width"],
                               textvariable=self.string_var,
                               name=kwargs["entry"]["name"])
        self.entry.bind("<Button-3>", kwargs["entry"]["right_click"])
        if kwargs["entry"].get("callback"):
            self.entry.bind("<FocusOut>", entry_callback)
        self.entry.grid(row=row, column=1, sticky="w", padx=(0, 5), pady=(0, 5))

        # Create a Button
        def button_callback() -> str | None:
            filename = kwargs["button"]["command"](self.string_var, *kwargs["button"]["args"])
            if filename:
                if kwargs["entry"]["name"] == "input_entry":
                    kwargs["entry"]["IPTranslator"].input_file = filename
                elif kwargs["entry"]["name"] == "ref_entry":
                    kwargs["entry"]["IPTranslator"].ref_file = filename
            return filename

        self.button = ttk.Button(parent,
                                 text=kwargs["button"]["text"],
                                 command=button_callback,
                                 cursor="hand2",
                                 name=kwargs["button"]["name"])
        self.button.grid(row=row, column=2, sticky="w", padx=(0, 5), pady=(0, 5))

        # Create a Combobox
        self.checkbox_var = tk.IntVar()
        self.checkbox = ttk.Checkbutton(parent,
                                        variable=self.checkbox_var,
                                        onvalue=1, offvalue=0,
                                        takefocus=False,
                                        # command=kwargs["checkbox"]["command"],
                                        )
        self.checkbox.var = self.checkbox_var
        self.checkbox_var.set(kwargs["checkbox"]["default"])
        self.checkbox_var.trace_add("write", kwargs["checkbox"]["command"])
        self.checkbox.grid(row=row, column=3, sticky="w", padx=(0, 5), pady=(0, 5))
        if kwargs["checkbox"].get("state"):
            self.checkbox.config(state=kwargs["checkbox"]["state"], cursor=kwargs["checkbox"]["cursor"])


class StatusRow:
    def __init__(self, parent: ttk.Frame, row: int, **kwargs):
        """Create a row for each method in the application."""

        # Create  a Label
        self.label = ttk.Label(parent, text=kwargs["label"]["text"])
        self.label.grid(row=row, column=0, sticky="e", padx=(0, 5), pady=(0, 5))

        # Create a status Label
        self.status = ttk.Label(parent,
                                text=kwargs["status"]["text"],
                                foreground=kwargs["status"]["foreground"],
                                font=("Arial", 10, "bold"))
        self.status.grid(row=row, column=1, sticky="w", padx=(0, 5), pady=(0, 5))

        self.button = ttk.Button(parent,
                                 text=kwargs["button"]["text"],
                                 command=kwargs["button"]["command"],
                                 cursor="hand2")
        self.button.grid(row=row, column=2, sticky="w", padx=(0, 5), pady=(0, 5))

        # Create a Combobox
        self.checkbox_var = tk.IntVar()
        self.checkbox = ttk.Checkbutton(parent,
                                        variable=self.checkbox_var,
                                        onvalue=1, offvalue=0,
                                        takefocus=False,
                                        # command=kwargs["checkbox"]["command"],
                                        )
        self.checkbox.var = self.checkbox_var
        self.checkbox_var.set(kwargs["checkbox"]["default"])
        self.checkbox_var.trace_add("write", kwargs["checkbox"]["command"])
        self.checkbox.grid(row=row, column=3, sticky="w", padx=(0, 5), pady=(0, 5))
        if kwargs["checkbox"].get("state"):
            self.checkbox.config(state=kwargs["checkbox"]["state"], cursor=kwargs["checkbox"]["cursor"])


class CredentialsWindow:
    def __init__(self,
                 parent: tk.Tk | tk.Toplevel,
                 title: str,
                 app: str,
                 rows: list[list[str | bool]],
                 translator: IPTranslator,
                 icon: tk.PhotoImage,
                 **kwargs):
        """Create a window for credentials."""

        self.__parent = parent
        self.__title = title
        self.__icon = icon

        self.__app = app
        self.__rows = rows
        self.__translator = translator

        self.__window = tk.Toplevel(self.__parent)
        self.__window.title(self.__title)

        self.__right_click_menu = tk.Menu(self.__window, tearoff=0)
        self.right_click_fn()

        self.__window.focus_force()
        self.__window.grab_set()
        self.__window.focus_set()
        self.__window.transient(self.__parent)

        WIDTH = kwargs.get("width", 350)
        HEIGHT = kwargs.get("height", 230)

        self.__window.geometry(
            f'{WIDTH}x{HEIGHT}+{self.__parent.winfo_rootx() + 100}+{self.__parent.winfo_rooty() + 75}')
        self.__window.resizable(False, False)

        self.__frame = ttk.Frame(self.__window)
        self.__frame.pack(fill="both", expand=True)

        self.__title_label = ttk.Label(self.__frame, text=self.__title, font=("Arial", 10, "bold"))
        self.__title_label.grid(row=0, column=0, columnspan=2, padx=5, pady=5)

        # row = ["text", "default", "bold"]
        self.__rows_contents = {}
        for i, row in enumerate(self.__rows):
            self.__rows_contents[row[0]] = {"label": ttk.Label(), "entry": ttk.Entry(), "entry_var": tk.StringVar()}
            self.__rows_contents[row[0]]["label"] = ttk.Label(self.__frame,
                                                              text=row[0],
                                                              font=("Arial", 8, "bold" if row[2] else "normal"))
            self.__rows_contents[row[0]]["label"].grid(row=i + 1, column=0, sticky="e", padx=5, pady=5)

            self.__rows_contents[row[0]]["entry"] = ttk.Entry(self.__frame,
                                                              textvariable=self.__rows_contents[row[0]]["entry_var"],
                                                              width=37)
            self.__rows_contents[row[0]]["entry"].bind("<Button-3>", self.right_click)
            if "password" in row[0].lower():
                self.__rows_contents[row[0]]["entry"].config(show="*")
            self.__rows_contents[row[0]]["entry_var"].set(row[1])
            self.__rows_contents[row[0]]["entry"].grid(row=i + 1, column=1, sticky="w", padx=5, pady=5)

        self.__save_var = tk.IntVar()
        self.__save_checkbox = ttk.Checkbutton(self.__frame,
                                               text="Remember Me?",
                                               variable=self.__save_var,
                                               onvalue=1, offvalue=0,
                                               cursor="hand2",
                                               takefocus=False,
                                               width=15)
        self.__save_checkbox.grid(row=len(self.__rows) + 1, column=0, columnspan=2, sticky="w", padx=25, pady=5)

        self.__buttons_frame = ttk.Frame(self.__frame)
        self.__buttons_frame.grid(row=len(self.__rows) + 2, column=0, columnspan=2, padx=5, pady=5)

        self.__save_button = ttk.Button(self.__buttons_frame,
                                        text="Save",
                                        command=self.save,
                                        cursor="hand2",
                                        width=15)
        self.__save_button.grid(row=0, column=0, sticky="w", padx=(20, 5), pady=5)

        self.__cancel_button = ttk.Button(self.__buttons_frame,
                                          text="Cancel",
                                          command=self.destroy,
                                          cursor="hand2",
                                          width=15)
        self.__cancel_button.grid(row=0, column=1, sticky="e", padx=5, pady=5)

    def right_click_fn(self) -> None:

        # Create a right click functions
        def cut():
            widget = self.__window.focus_get()
            if isinstance(widget, tk.Entry):
                widget.event_generate("<<Cut>>")

        def copy():
            widget = self.__window.focus_get()
            if isinstance(widget, tk.Entry):
                widget.event_generate("<<Copy>>")

        def paste():
            widget = self.__window.focus_get()
            if isinstance(widget, tk.Entry):
                widget.event_generate("<<Paste>>")

        def delete():
            widget = self.__window.focus_get()
            if isinstance(widget, tk.Entry):
                widget.event_generate("<Delete>")

        def select_all():
            widget = self.__window.focus_get()
            if isinstance(widget, tk.Entry):
                widget.event_generate("<<SelectAll>>")

        def clear_all():
            widget = self.__window.focus_get()
            if isinstance(widget, tk.Entry):
                widget.delete(0, "end")

        # Create a right click menu
        self.__right_click_menu = tk.Menu(self.__window, tearoff=0)
        self.__right_click_menu.add_command(label="Cut", command=cut)
        self.__right_click_menu.add_command(label="Copy", command=copy)
        self.__right_click_menu.add_command(label="Paste", command=paste)
        self.__right_click_menu.add_command(label="Delete", command=delete)
        self.__right_click_menu.add_separator()
        self.__right_click_menu.add_command(label="Select All", command=select_all)
        self.__right_click_menu.add_command(label="Clear All", command=clear_all)

    def right_click(self, event) -> None:
        event.widget.focus()
        self.__right_click_menu.tk_popup(event.x_root, event.y_root)

    def save(self, remember: int = 0) -> bool | None:
        self.__save_button.config(state="disabled", cursor="arrow")
        self.__cancel_button.config(state="disabled", cursor="arrow")
        self.__window.protocol("WM_DELETE_WINDOW", lambda: False)

        __strings = []
        remember = self.__save_var.get()
        for key, value in self.__rows_contents.items():
            __strings.append(value["entry_var"].get().strip())

        if self.__app in ["pan", "forti", "apic"]:
            __connected = None
            if self.__app in ["pan", "apic"]:
                __ip = __strings[0]
                __username = __strings[1]
                __password = __strings[2]
                __vsys = __strings[3]

                if self.__app == "apic" and not __vsys:
                    messagebox.showerror("Error",
                                         "Class(es) can not be empty.",
                                         parent=self.__window)
                    self.__save_button.config(state="normal", cursor="hand2")
                    self.__cancel_button.config(state="normal", cursor="hand2")
                    self.__window.protocol("WM_DELETE_WINDOW", self.destroy)
                    return False

                if self.__app == "pan":
                    __connected = PropagatingThread(target=self.__translator.connect_pan,
                                                    kwargs={"ip": __ip,
                                                            "username": __username,
                                                            "password": __password,
                                                            "vsys": __vsys,
                                                            "window": self.__window},
                                                    daemon=True,
                                                    name="connect_pan")
                elif self.__app == "apic":
                    __connected = PropagatingThread(target=self.__translator.connect_apic,
                                                    kwargs={"ip": __ip,
                                                            "username": __username,
                                                            "password": __password,
                                                            "apic_class": __vsys,
                                                            "window": self.__window},
                                                    daemon=True,
                                                    name="connect_apic")

            elif self.__app == "forti":
                __ip = __strings[0]
                __port = __strings[1]
                __username = __strings[2]
                __password = __strings[3]
                __vdom = __strings[4]

                __connected = PropagatingThread(target=self.__translator.connect_forti,
                                                kwargs={"ip": __ip,
                                                        "port": __port,
                                                        "username": __username,
                                                        "password": __password,
                                                        "vdom": __vdom,
                                                        "window": self.__window},
                                                daemon=True,
                                                name="connect_forti")

            __connected.start()
            while __connected.is_alive():
                self.__window.update()
                self.__parent.update()
                time.sleep(0.1)
            __connected = __connected.join()

            if not __connected:
                if not tk.messagebox.askyesno("Warning",
                                              "Connection failed, do you want to save the credentials anyway?",
                                              icon="warning",
                                              default="no",
                                              parent=self.__window):
                    self.__save_button.config(state="normal", cursor="hand2")
                    self.__cancel_button.config(state="normal", cursor="hand2")
                    self.__window.protocol("WM_DELETE_WINDOW", self.destroy)
                    return False

            if self.__app == "pan":
                self.__translator.pan_ip = __ip
                self.__translator.pan_username = __username
                self.__translator.pan_password = __password
                self.__translator.pan_vsys = __vsys

                if remember:
                    self.__translator.save_credentials(app="pan")

            elif self.__app == "forti":
                self.__translator.forti_ip = __ip
                self.__translator.forti_port = __port
                self.__translator.forti_username = __username
                self.__translator.forti_password = __password
                self.__translator.forti_vdom = __vdom

                if remember:
                    self.__translator.save_credentials(app="forti")

            elif self.__app == "apic":
                self.__translator.apic_ip = __ip
                self.__translator.apic_username = __username
                self.__translator.apic_password = __password
                self.__translator.apic_class = __vsys

                if remember:
                    self.__translator.save_credentials(app="apic")

        elif self.__app == "dns":
            __servers = __strings.copy()

            for i in range(3, -1, -1):
                if not __servers[i]:
                    __servers.pop(i)
            if len(__servers) < 4:
                __servers.extend([""] * (4 - len(__servers)))

            if not __servers:
                messagebox.showerror("Error",
                                     "DNS servers can not be empty.",
                                     parent=self.__window)
                self.__save_button.config(state="normal", cursor="hand2")
                self.__cancel_button.config(state="normal", cursor="hand2")
                self.__window.protocol("WM_DELETE_WINDOW", self.destroy)
                return False

            __connected = PropagatingThread(target=self.__translator.check_dns_servers,
                                            args=(__servers,),
                                            daemon=True,
                                            name="check_dns_servers")
            __connected.start()
            while __connected.is_alive():
                self.__window.update()
                self.__parent.update()
                time.sleep(0.1)
            __good_servers = __connected.join()

            __bad_servers = []
            for server in __servers:
                if server and (server not in __good_servers):
                    __bad_servers.append(server)

            if not len(__good_servers):
                if not messagebox.askyesno("Error",
                                           "No reachable DNS servers found, do you want to save these servers anyway?",
                                           icon="error",
                                           parent=self.__window):
                    self.__save_button.config(state="normal", cursor="hand2")
                    self.__cancel_button.config(state="normal", cursor="hand2")
                    self.__window.protocol("WM_DELETE_WINDOW", self.destroy)
                    return False

            elif len(__bad_servers):
                if not messagebox.askyesno("Warning",
                                           f"Invalid DNS servers found:\n    {'\n    '.join(__bad_servers)}\n\nDo you want to save the servers anyway?",
                                           icon="warning",
                                           parent=self.__window):
                    self.__save_button.config(state="normal", cursor="hand2")
                    self.__cancel_button.config(state="normal", cursor="hand2")
                    self.__window.protocol("WM_DELETE_WINDOW", self.destroy)
                    return False

            self.__translator.dns_servers = __servers.copy()
            self.__translator.resolvers = __good_servers.copy()

            if remember:
                self.__translator.save_credentials(app="dns")

        self.destroy()

    def destroy(self) -> None:
        self.__window.destroy()
        self.__parent.attributes('-disabled', False)
        self.__parent.update()
        self.__parent.focus_force()

    def mainloop(self) -> None:
        if self.__app == "pan":
            if all([self.__translator.pan_ip == "",
                    self.__translator.pan_username == "",
                    self.__translator.pan_password == ""]):
                self.__translator.import_credentials("pan")

        elif self.__app == "forti":
            if all([self.__translator.forti_ip == "",
                    self.__translator.forti_username == "",
                    self.__translator.forti_password == ""]):
                self.__translator.import_credentials("forti")

        elif self.__app == "apic":
            if all([self.__translator.apic_ip == "",
                    self.__translator.apic_username == "",
                    self.__translator.apic_password == ""]):
                self.__translator.import_credentials("apic")

        elif self.__app == "dns":
            if not self.__translator.dns_servers:
                self.__translator.import_credentials("dns")

        self.__parent.attributes('-disabled', True)
        self.__window.protocol("WM_DELETE_WINDOW", self.destroy)
        self.__window.bind('<Escape>', self.destroy)
        self.__window.bind('<Return>', self.save)
        self.__window.iconphoto(False, self.__icon)
        self.__window.mainloop()


class GUI:
    """A class that represent the windows gui of the application."""

    def __init__(self,
                 name: str = "IP Address Translator",
                 short_name: str = "IAT",
                 version: str = "1.0",
                 width: int = 550,
                 height: int = 490,
                 resizable: tuple[bool, bool] = (False, False),
                 icon_data: str = utilities.ICON_DATA,
                 *args, **kwargs):
        """Initialize the GUI class."""

        self.__name: str = name
        self.__short_name: str = short_name
        self.__version: str = version

        self.__width: int = width
        self.__height: int = height
        self.__resizable: tuple[bool, bool] = resizable
        self.__args = args
        self.__kwargs = kwargs

        self.__author: str = "xAbdalla"
        self.__github_link: str = "https://github.com/xAbdalla"

        self.__full_name: str = f"{self.__name} ({self.__short_name})"
        self.__title: str = f"{self.__full_name} v{self.__version} by {self.__author}"
        self.__title_label: str = f"{self.__name} v{self.__version}"

        self.__root: tk.Tk = tk.Tk()
        self.__root.title(self.__title)

        self.__icon: tk.PhotoImage = tk.PhotoImage(data=b64decode(icon_data))
        self.__github_icon: tk.PhotoImage = tk.PhotoImage(data=b64decode(utilities.GITHUB_ICON_DATA))
        self.__github_icon = self.__github_icon.zoom(1).subsample(18)

        self.__info_icon: tk.PhotoImage = tk.PhotoImage(data=b64decode(utilities.INFO_ICON_DATA))
        self.__info_icon = self.__info_icon.zoom(2).subsample(13)

        self.__SCREEN_WIDTH = self.__root.winfo_screenwidth()
        self.__SCREEN_HEIGHT = self.__root.winfo_screenheight()

        self.__right_click_menu = tk.Menu(self.__root, tearoff=0)
        self.right_click_fn()

        self.log = tk.StringVar()
        if self.__kwargs.get("log"):
            self.log.trace_add("write", self.write_logs)

            # Check if the log file exists and big delete it
            # if os.path.exists("logs.txt"):
            #     if os.path.getsize("logs.txt") > int(5 * 1024 * 1024):
            #         os.remove("logs.txt")
        self.log.set("INFO - Application started.")

        self.IPTranslator: IPTranslator = IPTranslator()
        self.IPTranslator.parent = self
        self.__start_thread: PropagatingThread | None = None
        self.__start_flag: bool = False
        self.__stop_flag: bool = False
        self.__methods_flags: list[bool] = [False for _ in range(5)]

        self.info_type_label = None
        self.msg_text = tk.StringVar()
        self.info_msg_label = None
        self.input_row = None
        self.ref_row = None
        self.pan_row = None
        self.forti_row = None
        self.apic_row = None
        self.dns_row = None
        self.progress_bar = None
        self.progress_percentage = None
        self.start_button = None

        self.init_root_window()

    #######################################################

    def start_status(self) -> bool:
        return self.__start_flag

    def stop_status(self) -> bool:
        return self.__stop_flag

    def get_root(self) -> tk.Tk:
        return self.__root

    def get_methods_flags(self) -> list[bool]:
        return self.__methods_flags

    #######################################################

    def right_click_fn(self) -> None:

        # Create a right click functions
        def cut():
            widget = self.__root.focus_get()
            if isinstance(widget, tk.Entry):
                widget.event_generate("<<Cut>>")

        def copy():
            widget = self.__root.focus_get()
            if isinstance(widget, tk.Entry):
                widget.event_generate("<<Copy>>")

        def paste():
            widget = self.__root.focus_get()
            if isinstance(widget, tk.Entry):
                widget.event_generate("<<Paste>>")

        def delete():
            widget = self.__root.focus_get()
            if isinstance(widget, tk.Entry):
                widget.event_generate("<Delete>")

        def select_all():
            widget = self.__root.focus_get()
            if isinstance(widget, tk.Entry):
                widget.event_generate("<<SelectAll>>")

        def clear_all():
            widget = self.__root.focus_get()
            if isinstance(widget, tk.Entry):
                widget.delete(0, "end")

        # Create a right click menu
        self.__right_click_menu = tk.Menu(self.__root, tearoff=0)
        self.__right_click_menu.add_command(label="Cut", command=cut)
        self.__right_click_menu.add_command(label="Copy", command=copy)
        self.__right_click_menu.add_command(label="Paste", command=paste)
        self.__right_click_menu.add_command(label="Delete", command=delete)
        self.__right_click_menu.add_separator()
        self.__right_click_menu.add_command(label="Select All", command=select_all)
        self.__right_click_menu.add_command(label="Clear All", command=clear_all)

    def right_click(self, event) -> None:
        event.widget.focus()
        self.__right_click_menu.tk_popup(event.x_root, event.y_root)

    def write_logs(self, *arg):
        log = self.log.get()
        with open("logs.txt", "a") as f:
            if log:
                f.write(f"{datetime.now().strftime("%Y-%m-%d %H:%M:%S")} - {log}\n")
            else:
                f.write(f"{"=" * 50}\n")

    def disable_all_buttons(self) -> None:
        self.__root.protocol("WM_DELETE_WINDOW", self.stop)

        self.input_row.entry.config(state="disabled", cursor="arrow")
        self.input_row.button.config(state="disabled", cursor="arrow")

        self.ref_row.entry.config(state="disabled", cursor="arrow")
        self.ref_row.button.config(state="disabled", cursor="arrow")
        self.ref_row.checkbox.config(state="disabled", cursor="arrow")

        self.pan_row.button.config(state="disabled", cursor="arrow")
        self.pan_row.checkbox.config(state="disabled", cursor="arrow")

        self.forti_row.button.config(state="disabled", cursor="arrow")
        self.forti_row.checkbox.config(state="disabled", cursor="arrow")

        self.apic_row.button.config(state="disabled", cursor="arrow")
        self.apic_row.checkbox.config(state="disabled", cursor="arrow")

        self.dns_row.button.config(state="disabled", cursor="arrow")
        self.dns_row.checkbox.config(state="disabled", cursor="arrow")

    def enable_all_buttons(self) -> None:
        self.__root.protocol("WM_DELETE_WINDOW", self.destroy)

        self.input_row.entry.config(state="normal", cursor="xterm")
        self.input_row.button.config(state="normal", cursor="hand2")

        self.ref_row.entry.config(state="normal", cursor="xterm")
        self.ref_row.button.config(state="normal", cursor="hand2")
        self.ref_row.checkbox.config(state="normal", cursor="hand2")

        self.pan_row.button.config(state="normal", cursor="hand2")
        self.pan_row.checkbox.config(state="normal", cursor="hand2")

        self.forti_row.button.config(state="normal", cursor="hand2")
        self.forti_row.checkbox.config(state="normal", cursor="hand2")

        self.apic_row.button.config(state="normal", cursor="hand2")
        self.apic_row.checkbox.config(state="normal", cursor="hand2")

        self.dns_row.button.config(state="normal", cursor="hand2")
        self.dns_row.checkbox.config(state="normal", cursor="hand2")

    def disable_connect_buttons(self) -> None:
        self.pan_row.button.config(state="disabled", cursor="arrow")
        self.forti_row.button.config(state="disabled", cursor="arrow")
        self.apic_row.button.config(state="disabled", cursor="arrow")
        self.dns_row.button.config(state="disabled", cursor="arrow")
        self.start_button.config(state="disabled", cursor="arrow")

    def enable_connect_buttons(self) -> None:
        __pan = self.pan_row.button["text"] == "Connect"
        __forti = self.forti_row.button["text"] == "Connect"
        __apic = self.apic_row.button["text"] == "Connect"
        __dns = self.dns_row.button["text"] == "Check"

        if all([__pan, __forti, __apic, __dns]):
            self.pan_row.button.config(state="normal", cursor="hand2")
            self.forti_row.button.config(state="normal", cursor="hand2")
            self.apic_row.button.config(state="normal", cursor="hand2")
            self.dns_row.button.config(state="normal", cursor="hand2")
            if True in self.__methods_flags:
                self.start_button.config(state="normal", cursor="hand2")
        else:
            self.disable_connect_buttons()

    def set_info_status(self, status: str, foreground: str, cursor: str) -> None:
        """Set the status of the info label."""
        self.info_type_label.config(text=status, foreground=foreground, cursor=cursor)

    def set_info_message(self, message: str, foreground: str, cursor: str) -> None:
        """Set the message of the info label."""
        self.msg_text.set(message)
        self.info_msg_label.config(foreground=foreground, cursor=cursor)

    def check_methods(self, event=None, *args) -> bool:
        if self.ref_row.checkbox_var.get():
            self.__methods_flags[0] = True
        else:
            self.__methods_flags[0] = False

        if self.pan_row.checkbox_var.get():
            self.__methods_flags[1] = True
        else:
            self.__methods_flags[1] = False

        if self.forti_row.checkbox_var.get():
            self.__methods_flags[2] = True
        else:
            self.__methods_flags[2] = False

        if self.apic_row.checkbox_var.get():
            self.__methods_flags[3] = True
        else:
            self.__methods_flags[3] = False

        if self.dns_row.checkbox_var.get():
            self.__methods_flags[4] = True
        else:
            self.__methods_flags[4] = False

        if True in self.__methods_flags:
            if not self.__start_flag:
                self.set_info_status("INFO", "green", "arrow")
                self.set_info_message("Click the (Start) button to start the translation process",
                                      "green", "arrow")

            if (str(self.pan_row.button["state"]) == "normal" and
                    str(self.forti_row.button["state"]) == "normal" and
                    str(self.apic_row.button["state"]) == "normal" and
                    str(self.dns_row.button["state"]) == "normal"):
                self.start_button.config(state="normal", cursor="hand2")
                self.__root.bind("<Return>", self.start)
            return True

        if not self.__start_flag:
            self.set_info_status("INFO", "green", "arrow")
            self.set_info_message("Please select one searching method at least",
                                  "red", "arrow")
        self.start_button.config(state="disabled", cursor="arrow")
        self.__root.unbind("<Return>")
        return False

    #######################################################

    def info_window(self) -> None:
        def destroy_info_window(event=None):
            info_window.destroy()
            self.__root.attributes('-disabled', False)
            self.__root.update()
            self.__root.focus_force()

        font_bold = ("Arial", 10, "bold")
        font_normal = ("Arial", 10)

        objective_text = f"""{self.__name} application assists you in mapping IPs from network logs to descriptive \
object names or domain names, making log analysis more straightforward."""
        features_text = f"""➞ Various Input Options:
      • Accepts Excel/CSV/Text files.
      • Direct input of single IP address, subnet, range, or list (comma separated).

➞ Files Specifications:
      • Input Text files must have one subnet per line.
      • Input CSV files must have a column named 'Subnet'.
      • Input Excel files could have multiple sheets, each must containing a column named 'Subnet'.
      • Reference CSV files require three columns: 'Tenant', 'Address object', and 'Subnet'.
      • Reference Excel files could have multiple sheets, each must containing the three columns.

➞ Various Searching Methods:
      • Reference File: Searches for matches in a user-provided reference file.
      • Palo Alto: Connects via SSH to Panorama to import address objects.
      • Fortinet: Connects via REST API to FortiGate to import address objects.
      • Cisco ACI: Connects via SSH to APIC to import address objects based on the specified Class.
      • Reverse DNS: Resolves IPs to domain names using the reverse DNS servers.

➞ Palo Alto Panorama/FW Specifications:
      • Ensure Panorama/Firewall is reachable and has CLI access.
      • Leave "VSYS" field empty if you want to retrieve address objects from all virtual systems.
      • The program may fail to import the addresses due to slow response, just try again until it works.

➞ Fortinet FortiGate Specifications:
      • Ensure FortiGate is reachable and you have REST API enabled.
      • Leave "VDOM" field empty if you want to retrieve address objects from all virtual domains.

➞ Cisco ACI Specifications:
      • Ensure APIC is reachable and has CLI access.
      • Specify the Class of the address objects to be searched.
      • The program searches the "dn" attribute exclusively.

➞ DNS Resolver:
      • Resolves IPs to domain names using the system dns servers or a user-provided DNS servers.
      • You can specify up to four DNS servers.

➞ Private Information Storage:
      • An option to save your credentials for future use and avoid re-entering them.
      • Stored information are encrypted.
      • Credentials are stored locally in the application directory.

➞ Encrypted Credentials Storage:
      • An option to save your credentials for future use and avoid re-entering them.
      • Credentials are stored locally in the application directory.
      • Stored information are encrypted for security purposes.
      • The encryption key is unique to each user and machine.

➞ Saving the Output:
      • Results are exported to a new Excel (.xlsx) file for ease of access and analysis.
      • The generated file contains the original data along with the mapped names.
      • The user can specify the output file name and location to prevent overwriting.

➞ Logging
      • Detailed logs are generated for each operation.
      • Logs are saved in a separate file for future reference.
      • Logs are also displayed in the GUI for immediate feedback.
      • Logs are color-coded for better readability.

➞ Error Handling
      • Detailed error messages are displayed in the GUI.
      • Logs are generated for each error for future reference.
      • Errors are color-coded for better readability."""
        usage_text = f"""• Fill in the required fields and select the desired search method.
• Ensure the chosen search methods are accessible and correctly configured.
• Provide necessary credentials and remember to save them if needed.
• For Cisco ACI, you must specify the correct Class for targeted searches.
• Review the generated Excel file for mapped IPs based on the selected search methods.
• Review the logs for detailed information about the operation.
• For any issues or inquiries, please contact the author."""

        info_window = tk.Toplevel(self.__root)
        info_window.title("Help")
        info_window.focus_force()
        info_window.grab_set()
        info_window.focus_set()
        info_window.transient(self.__root)
        self.__root.attributes('-disabled', True)

        info_window.geometry(f'{665}x{580}+{self.__root.winfo_rootx() - 65}+{self.__root.winfo_rooty() - 100}')
        info_window.resizable(False, False)
        info_window.iconphoto(False, self.__info_icon)

        info_frame = VerticalScrolledFrame(info_window)
        info_frame.pack(fill="both", expand=True)

        # Objective
        objective_label = ttk.Label(info_frame.interior, text="Objective:", font=font_bold)
        objective_label.pack(pady=(10, 1), padx=10, anchor="w")

        objective_text = ttk.Label(info_frame.interior, text=objective_text, font=font_normal, justify="left",
                                   wraplength=620)
        objective_text.pack(pady=0, padx=20, anchor="w")

        # Features
        features_label = ttk.Label(info_frame.interior, text="Features:", font=font_bold)
        features_label.pack(pady=(10, 1), padx=10, anchor="w")

        features_text = ttk.Label(info_frame.interior, text=features_text, font=font_normal, justify="left",
                                  wraplength=620)
        features_text.pack(pady=0, padx=20, anchor="w")

        # Usage
        usage_label = ttk.Label(info_frame.interior, text="Usage:", font=font_bold)
        usage_label.pack(pady=(10, 1), padx=10, anchor="w")

        usage_text = ttk.Label(info_frame.interior, text=usage_text, font=font_normal, justify="left",
                               wraplength=620)
        usage_text.pack(pady=0, padx=20, anchor="w")

        author_label = ttk.Label(info_frame.interior, text=f"{self.__title}", font=font_bold)
        author_label.pack(pady=(10, 0), padx=10, anchor="w")

        feedback_label = ttk.Label(info_frame.interior, text=f"Feel free to provide feedback or report issues.",
                                   font=font_normal)
        feedback_label.pack(pady=0, padx=10, anchor="w")

        github_label = ttk.Label(info_frame.interior, text=f"GitHub",
                                 font=font_bold, cursor="hand2", foreground="blue")
        github_label.pack(pady=(0, 5), padx=10, anchor="w")
        github_label.bind("<Button-1>",
                          lambda e: webbrowser.open_new(f"{self.__github_link}/{self.__name.replace(" ", "_")}"))

        s2 = ttk.Style()
        s2.configure("close.TButton", font=font_bold)

        close_button = ttk.Button(info_frame.interior, text="Close", style="close.TButton",
                                  command=destroy_info_window, cursor="hand2")
        close_button.pack(pady=(10, 5), padx=10, anchor="center", ipady=5, ipadx=75)

        info_window.bind('<Escape>', destroy_info_window)
        info_window.protocol("WM_DELETE_WINDOW", destroy_info_window)

        info_window.mainloop()

    #######################################################

    def get_pan_credentials(self) -> None:
        if all([self.IPTranslator.pan_ip == "",
                self.IPTranslator.pan_username == "",
                self.IPTranslator.pan_password == "",
                ]):
            self.IPTranslator.import_credentials("pan")

        __rows = [
            ["PAN IP:", self.IPTranslator.pan_ip, True],
            ["Username:", self.IPTranslator.pan_username, True],
            ["Password:", self.IPTranslator.pan_password, True],
            ["VSYS:", self.IPTranslator.pan_vsys, False]
        ]

        __pan_window = CredentialsWindow(parent=self.__root,
                                         title="PAN Credentials",
                                         app="pan",
                                         rows=__rows,
                                         translator=self.IPTranslator,
                                         icon=self.__icon,
                                         width=320,
                                         height=230
                                         )
        __pan_window.mainloop()

    #######################################################

    def get_forti_credentials(self) -> None:
        if all([
            self.IPTranslator.forti_ip == "",
            self.IPTranslator.forti_username == "",
            self.IPTranslator.forti_password == "",
        ]):
            self.IPTranslator.import_credentials("forti")

        __rows = [
            ["Forti IP:", self.IPTranslator.forti_ip, True],
            ["Port:", self.IPTranslator.forti_port, True],
            ["Username:", self.IPTranslator.forti_username, True],
            ["Password:", self.IPTranslator.forti_password, True],
            ["VDOM:", self.IPTranslator.forti_vdom, False]
        ]

        __forti_window = CredentialsWindow(parent=self.__root,
                                           title="Forti Credentials",
                                           app="forti",
                                           rows=__rows,
                                           translator=self.IPTranslator,
                                           icon=self.__icon,
                                           width=320,
                                           height=260
                                           )
        __forti_window.mainloop()

    #######################################################

    def get_apic_credentials(self) -> None:
        if all([self.IPTranslator.apic_ip == "",
                self.IPTranslator.apic_username == "",
                self.IPTranslator.apic_password == "",
                self.IPTranslator.apic_class == "",
                ]):
            self.IPTranslator.import_credentials("apic")

        __rows = [
            ["APIC IP:", self.IPTranslator.apic_ip, True],
            ["Username:", self.IPTranslator.apic_username, True],
            ["Password:", self.IPTranslator.apic_password, True],
            ["Class:", self.IPTranslator.apic_class, True]
        ]

        __apic_window = CredentialsWindow(parent=self.__root,
                                          title="APIC Credentials",
                                          app="apic",
                                          rows=__rows,
                                          translator=self.IPTranslator,
                                          icon=self.__icon,
                                          width=320,
                                          height=230
                                          )
        __apic_window.mainloop()

    #######################################################

    def get_dns_servers(self) -> None:
        if not self.IPTranslator.dns_servers:
            self.IPTranslator.import_credentials("dns")

        __rows = [
            ["Server 1:", self.IPTranslator.dns_servers[0], True],
            ["Server 2:", self.IPTranslator.dns_servers[1], False],
            ["Server 3:", self.IPTranslator.dns_servers[2], False],
            ["Server 4:", self.IPTranslator.dns_servers[3], False]
        ]

        __dns_window = CredentialsWindow(parent=self.__root,
                                         title="DNS Servers",
                                         app="dns",
                                         rows=__rows,
                                         translator=self.IPTranslator,
                                         icon=self.__icon,
                                         width=305,
                                         height=230
                                         )
        __dns_window.mainloop()

    #######################################################

    def pre_start(self) -> None:
        self.post_start()
        self.__start_flag = True
        self.__stop_flag = False
        self.IPTranslator.__stop = False

        s = ttk.Style()
        s.configure("start.TButton", font=("Arial bold", 15), foreground="red")

        self.start_button.config(text="Stop", state="normal", style="start.TButton", command=self.stop, cursor="hand2")
        self.__root.bind("<Return>", self.stop)

        self.disable_all_buttons()

    def start(self, event=None) -> bool:
        self.pre_start()
        self.start_button.config(command=lambda: False)
        self.log.set("INFO - Starting the translation process ...")

        # Check Input File
        if not self.IPTranslator.inputs:
            __input_thread = PropagatingThread(target=self.IPTranslator.check_input_file,
                                               daemon=True)
            __input_thread.start()
            while __input_thread.is_alive():
                self.__root.update()
                time.sleep(0.1)
            __input = __input_thread.join()
            if not __input:
                self.set_info_status("ERROR", "red", "arrow")
                self.set_info_message("Input file is missing or invalid.", "red", "arrow")
                self.log.set("ERROR - Input file is missing or invalid.")
                self.post_start()
                return False

        if self.IPTranslator.inputs:
            if not self.IPTranslator.inputs_no:
                self.set_info_status("ERROR", "red", "arrow")
                self.set_info_message("No valid IP addresses found in the input file.", "red", "arrow")
                self.log.set(f"ERROR - No valid IP addresses found in the input file ({self.IPTranslator.input_file}).")
                self.post_start()
                return False

            self.set_info_status("INFO", "green", "arrow")
            self.set_info_message(f"Input file is valid. {self.IPTranslator.all_inputs_no} IP addresses found.",
                                  "green", "arrow")
            self.log.set(f"INFO - Input file is valid. {self.IPTranslator.all_inputs_no} IP addresses found.")

        # Check Used Methods
        self.check_methods()

        if not (True in self.__methods_flags):
            self.post_start()
            self.set_info_status("ERROR", "red", "arrow")
            self.set_info_message("Please select one searching method at least", "red", "arrow")
            return False

        if self.__methods_flags[0]:
            # Check Reference File
            if not self.IPTranslator.refs:
                __ref_thread = PropagatingThread(target=self.IPTranslator.check_ref_file,
                                                 daemon=True)
                __ref_thread.start()
                while __ref_thread.is_alive():
                    self.__root.update()
                    time.sleep(0.1)
                __ref = __ref_thread.join()
                if not __ref:
                    self.post_start()
                    self.ref_row.checkbox_var.set(0)

                    self.set_info_status("ERROR", "red", "arrow")
                    self.set_info_message("Reference file is missing or invalid.", "red", "arrow")
                    self.log.set(
                        f"ERROR - Reference file is missing or invalid ({self.IPTranslator.ref_file}).")
                    return False

            if self.IPTranslator.refs:
                if not self.IPTranslator.ref_no:
                    self.post_start()
                    self.ref_row.checkbox_var.set(0)

                    self.set_info_status("ERROR", "red", "arrow")
                    self.set_info_message("No valid addresses found in the reference file.", "red", "arrow")
                    self.log.set(
                        f"ERROR - No valid addresses found in the reference file ({self.IPTranslator.ref_file}).")
                    return False

                self.set_info_status("INFO", "green", "arrow")
                self.set_info_message(f"Reference file is valid. {self.IPTranslator.ref_no} addresses found.",
                                      "green", "arrow")
                self.log.set(f"INFO - Reference file is valid. {self.IPTranslator.ref_no} addresses found.")

        if self.__methods_flags[1]:
            if not self.IPTranslator.pan_addresses:
                # Connect to Palo Alto
                __connect_pan_thread = PropagatingThread(target=self.IPTranslator.connect_pan,
                                                         daemon=True)
                __connect_pan_thread.start()
                while __connect_pan_thread.is_alive():
                    self.__root.update()
                    time.sleep(0.1)
                __pan = __connect_pan_thread.join()
                if not __pan:
                    self.post_start()
                    self.pan_row.checkbox_var.set(0)

                    self.set_info_status("ERROR", "red", "arrow")
                    self.set_info_message("Failed to connect to PAN.", "red", "arrow")
                    self.log.set("ERROR - Failed to connect to PAN.")
                    return False

                else:
                    self.pre_start()
                    __import_pan_thread = PropagatingThread(target=self.IPTranslator.import_pan,
                                                            daemon=True)
                    __import_pan_thread.start()
                    while __import_pan_thread.is_alive():
                        self.__root.update()
                        time.sleep(0.1)
                    __import_pan = __import_pan_thread.join()
                    if not __import_pan:
                        self.post_start()
                        self.pan_row.checkbox_var.set(0)

                        self.set_info_status("ERROR", "red", "arrow")
                        self.set_info_message("No addresses found in PAN.", "red", "arrow")
                        self.log.set("ERROR - No addresses found in PAN.")
                        return False
                    else:
                        pan_no = len(self.IPTranslator.pan_addresses)
                        self.set_info_status("INFO", "green", "arrow")
                        self.set_info_message(f"{pan_no} addresses found in PAN.",
                                              "green", "arrow")
                        self.log.set(f"INFO - {pan_no} addresses found in PAN.")

            else:
                pan_no = len(self.IPTranslator.pan_addresses)
                self.set_info_status("INFO", "green", "arrow")
                self.set_info_message(f"{pan_no} addresses found in PAN.",
                                      "green", "arrow")
                self.log.set(f"INFO - {pan_no} addresses found in PAN.")

        if self.__methods_flags[2]:
            if not self.IPTranslator.forti_addresses:
                # Connect to Fortinet
                __connect_forti_thread = PropagatingThread(target=self.IPTranslator.connect_forti,
                                                           daemon=True)
                __connect_forti_thread.start()
                while __connect_forti_thread.is_alive():
                    self.__root.update()
                    time.sleep(0.1)
                __forti = __connect_forti_thread.join()
                if not __forti:
                    self.post_start()
                    self.forti_row.checkbox_var.set(0)

                    self.set_info_status("ERROR", "red", "arrow")
                    self.set_info_message("Failed to connect to FortiGate.", "red", "arrow")
                    self.log.set("ERROR - Failed to connect to FortiGate.")
                    return False

                else:
                    self.pre_start()
                    __import_forti_thread = PropagatingThread(target=self.IPTranslator.import_forti,
                                                              daemon=True)
                    __import_forti_thread.start()
                    while __import_forti_thread.is_alive():
                        self.__root.update()
                        time.sleep(0.1)
                    __import_forti = __import_forti_thread.join()
                    if not __import_forti:
                        self.post_start()
                        self.forti_row.checkbox_var.set(0)

                        self.set_info_status("ERROR", "red", "arrow")
                        self.set_info_message("No addresses found in FortiGate.", "red", "arrow")
                        self.log.set("ERROR - No addresses found in FortiGate.")
                        return False
                    else:
                        forti_no = len(self.IPTranslator.forti_addresses)
                        self.set_info_status("INFO", "green", "arrow")
                        self.set_info_message(f"{forti_no} addresses found in FortiGate.",
                                              "green", "arrow")
                        self.log.set(f"INFO - {forti_no} addresses found in FortiGate.")

            else:
                forti_no = len(self.IPTranslator.forti_addresses)
                self.set_info_status("INFO", "green", "arrow")
                self.set_info_message(f"{forti_no} addresses found in FortiGate.",
                                      "green", "arrow")
                self.log.set(f"INFO - {forti_no} addresses found in FortiGate.")

        if self.__methods_flags[3]:
            if not self.IPTranslator.apic_addresses:
                # Connect to Cisco ACI
                __connect_apic_thread = PropagatingThread(target=self.IPTranslator.connect_apic,
                                                          daemon=True)
                __connect_apic_thread.start()
                while __connect_apic_thread.is_alive():
                    self.__root.update()
                    time.sleep(0.1)
                __apic = __connect_apic_thread.join()
                if not __apic:
                    self.post_start()
                    self.apic_row.checkbox_var.set(0)

                    self.set_info_status("ERROR", "red", "arrow")
                    self.set_info_message("Failed to connect to APIC.", "red", "arrow")
                    self.log.set("ERROR - Failed to connect to APIC.")
                    return False

                else:
                    self.pre_start()
                    __import_apic_thread = PropagatingThread(target=self.IPTranslator.import_apic,
                                                             daemon=True)
                    __import_apic_thread.start()
                    while __import_apic_thread.is_alive():
                        self.__root.update()
                        time.sleep(0.1)
                    __import_apic = __import_apic_thread.join()
                    if not __import_apic:
                        self.post_start()
                        self.apic_row.checkbox_var.set(0)

                        self.set_info_status("ERROR", "red", "arrow")
                        self.set_info_message("No addresses found in APIC.", "red", "arrow")
                        self.log.set("ERROR - No addresses found in APIC.")
                        return False
                    else:
                        apic_no = len(self.IPTranslator.apic_addresses)
                        self.set_info_status("INFO", "green", "arrow")
                        self.set_info_message(f"{apic_no} addresses found in APIC.",
                                              "green", "arrow")
                        self.log.set(f"INFO - {apic_no} addresses found in APIC.")

            else:
                apic_no = len(self.IPTranslator.apic_addresses)
                self.set_info_status("INFO", "green", "arrow")
                self.set_info_message(f"{apic_no} addresses found in APIC.",
                                      "green", "arrow")
                self.log.set(f"INFO - {apic_no} addresses found in APIC.")

        if self.__methods_flags[4]:
            if not self.IPTranslator.resolvers:
                # Get DNS Servers
                __dns_thread = PropagatingThread(target=self.IPTranslator.check_dns_servers,
                                                 daemon=True)
                __dns_thread.start()
                while __dns_thread.is_alive():
                    self.__root.update()
                    time.sleep(0.1)
                __resolvers = __dns_thread.join()
                if not __resolvers:
                    self.post_start()
                    self.dns_row.checkbox_var.set(0)

                    self.set_info_status("ERROR", "red", "arrow")
                    self.set_info_message("No reachable DNS servers found.", "red", "arrow")
                    self.log.set("ERROR - No reachable DNS servers found.")
                    return False

                else:
                    self.pre_start()
                    self.IPTranslator.dns_resolver = dns.resolver.Resolver()
                    self.IPTranslator.dns_resolver.nameservers = self.IPTranslator.resolvers

                    self.set_info_status("INFO", "green", "arrow")
                    self.set_info_message(f"{len(self.IPTranslator.resolvers)} DNS servers are reachable.",
                                          "green", "arrow")
                    self.log.set(f"INFO - {len(self.IPTranslator.resolvers)} DNS servers are reachable.")

            else:
                self.set_info_status("INFO", "green", "arrow")
                self.set_info_message(f"{len(self.IPTranslator.resolvers)} DNS servers are reachable.",
                                      "green", "arrow")
                self.log.set(f"INFO - {len(self.IPTranslator.resolvers)} DNS servers are reachable.")
                pass

        self.set_info_status("INFO", "green", "arrow")
        self.set_info_message("The Translation process is ready to start.", "green", "arrow")
        self.start_button.config(command=self.stop)

        invalid_records = ""

        if self.IPTranslator.input_invalids:
            invalid_records = f"Invalid Inputs: ({len(self.IPTranslator.input_invalids)} records)\n\n"
            for sheet, ip, row in self.IPTranslator.input_invalids:
                ip = re.sub("\n+", "[NL]", ip)
                invalid_records += f"Sheet: {sheet}\tRow: {row}\tIP: {ip}\n"
            invalid_records += "\n"

        if self.IPTranslator.ref_invalids:
            if invalid_records:
                invalid_records += "=" * 50 + "\n\n"
            invalid_records += f"Invalid References: ({len(self.IPTranslator.ref_invalids)} records)\n\n"
            for sheet, ip, row in self.IPTranslator.ref_invalids:
                ip = re.sub("\n+", "[NL]", ip)
                invalid_records += f"Sheet: {sheet}\tRow: {row}\tIP: {ip}\n"

        if invalid_records:
            with open("invalid_records.txt", "w") as f:
                if "[NL]" in invalid_records:
                    f.write("[NL] == New Line\n\n")
                f.write(invalid_records.strip())

        if len(self.IPTranslator.inputs.keys()) == 1:
            __sheet_str = "1 Sheet"
        else:
            __sheet_str = f"{len(self.IPTranslator.inputs.keys())} Sheets"

        __content_str = f"Input File Contents:\n{__sheet_str}, {self.IPTranslator.inputs_no} Valid Subnets, {len(self.IPTranslator.input_invalids)} Invalid Subnets."
        if self.IPTranslator.input_invalids:
            __content_str += f"\nInvalid details are saved in the 'invalid_records.txt' file.\n\n"
        else:
            __content_str += "\n\n"

        if self.__methods_flags[0]:
            if len(self.IPTranslator.refs.keys()) == 1:
                __sheet_str = "1 Sheet"
            else:
                __sheet_str = f"{len(self.IPTranslator.refs.keys())} Sheets"

            __content_str += f"Reference File Contents:\n{__sheet_str}, {self.IPTranslator.ref_no} Valid Subnets, {len(self.IPTranslator.ref_invalids)} Invalid Subnets."
            if self.IPTranslator.ref_invalids:
                __content_str += f"\nInvalid details are saved in the 'invalid_records.txt' file.\n\n"
            else:
                __content_str += "\n\n"

        if self.__methods_flags[1]:
            __content_str += f"Palo Alto: {len(self.IPTranslator.pan_addresses)} Address Objects.\n\n"

        if self.__methods_flags[2]:
            __content_str += f"FortiGate: {len(self.IPTranslator.forti_addresses)} Address Objects.\n\n"

        if self.__methods_flags[3]:
            __content_str += f"Cisco ACI: {len(self.IPTranslator.apic_addresses)} Address Objects.\n\n"

        if self.__methods_flags[4]:
            __content_str += f"DNS Servers: {len(self.IPTranslator.resolvers)} Servers.\n\n"

        ans = messagebox.askquestion("Info",
                                     f"{__content_str}\nClick 'Yes' to start the translation process.",
                                     parent=self.__root)

        if ans != "yes":
            self.set_info_status("INFO", "green", "arrow")
            self.set_info_message("The process has been canceled.", "red", "arrow")
            self.log.set("INFO - The process has been canceled by the user.")

            self.post_start()
            return False

        self.__start_thread = PropagatingThread(target=self.IPTranslator.Translate)
        self.__start_thread.start()

    def post_start(self) -> None:
        self.__start_flag = False
        self.__stop_flag = False
        self.IPTranslator.__stop = False

        try:
            if self.__start_thread.is_alive():
                self.__start_thread.stop()
        except:
            pass

        s = ttk.Style()
        s.configure("start.TButton", font=("Arial bold", 15), foreground="green")

        self.start_button.config(text="Start", state="normal", style="start.TButton", command=self.start,
                                 cursor="hand2")
        self.__root.bind("<Return>", self.start)

        self.enable_all_buttons()

    def stop(self, event=None) -> None:
        self.__stop_flag = True

        ans = messagebox.askyesnocancel("Warning",
                                        "Do you want to save the results before stopping the process?",
                                        icon="warning",
                                        default="yes",
                                        parent=self.__root)

        if ans is True:
            self.log.set("INFO - User chose to stop the process and save the results.")
            self.__start_flag = False

        elif ans is False:
            self.IPTranslator.__stop = True
            self.log.set("INFO - User chose to stop the process without saving the results.")

            if self.progress_bar['value'] >= self.progress_bar['maximum']:
                self.progress_percentage.config(text="100%", foreground="green")
            else:
                self.progress_percentage.config(foreground="red")

            self.post_start()
            self.IPTranslator.clear_var()
            self.__root.update()

            self.set_info_status("INFO", "green", "arrow")
            self.set_info_message("The process has been stopped successfully.", "red", "arrow")

        self.__stop_flag = False

    #######################################################

    def init_root_header(self) -> None:
        """Initialize the header of the application."""

        header = ttk.Label(self.__root, text=self.__name, font=("Arial", 20, "bold"))
        header.pack(pady=(15, 0), anchor="center")

        head_btn_frame = ttk.Frame(self.__root)
        head_btn_frame.pack(pady=(0, 10), padx=20, anchor="center")

        github_button = tk.Button(head_btn_frame, text=f" {self.__author}", font=("Arial bold", 12),
                                  image=self.__github_icon, borderwidth=0, relief="sunken", compound="left",
                                  cursor="hand2", command=lambda: webbrowser.open_new(self.__github_link),
                                  anchor="center")
        github_button.grid(row=0, column=0, padx=(5, 300), pady=(3, 5), sticky="w")

        info_button = tk.Button(head_btn_frame, text="Help ", font=("Arial bold", 12),
                                image=self.__info_icon, borderwidth=0, relief="sunken", compound="right",
                                cursor="hand2", command=lambda: self.info_window(), anchor="center")
        info_button.grid(row=0, column=1, padx=(0, 5), pady=(3, 5), sticky="e")

        separator(self.__root, orient="horizontal", fill="x", y_offset=(0, 10), x_offset=20)

    def init_root_info(self) -> None:
        """Initialize the info frame of the application."""

        info_frame = ttk.Frame(self.__root)
        info_frame.pack(pady=(0, 10), padx=20, anchor="center")

        self.info_type_label = ttk.Label(info_frame,
                                         text="INFO",
                                         font=("Arial", 10, "bold"),
                                         foreground="green",
                                         cursor="arrow")
        self.info_type_label.pack()

        def unbind_msg(*args, **kwargs):
            self.info_msg_label.unbind("<Button-1>")

        self.msg_text.set("Welcome to the IP Address Translator App!")
        self.info_msg_label = ttk.Label(info_frame,
                                        textvariable=self.msg_text,
                                        font=("Arial", 10, "bold"),
                                        foreground="green",
                                        cursor="arrow")
        self.msg_text.trace_add("write", unbind_msg)
        self.info_msg_label.pack()

        separator(self.__root, orient="horizontal", fill="x", y_offset=(0, 10), x_offset=20)

    def init_root_rows(self) -> None:
        """Initialize the rows of the application."""

        methods_frame = ttk.Frame(self.__root)
        methods_frame.pack(pady=(0, 10), padx=20, anchor="center")

        input_row_contents = {
            "label": {"text": "Input File / IP",
                      },
            "entry": {"width": 50,
                      "name": "input_entry",
                      "right_click": self.right_click,
                      "IPTranslator": self.IPTranslator,
                      "callback": self.IPTranslator.check_input_file
                      },
            "button": {"text": "Browse",
                       "name": "input_button",
                       "command": browse_file,
                       "args": (("xlsx", "csv", "txt", "xls", "xlsm", "xlsb", "odf", "ods", "odt"),
                                "Select an Input File")
                       },
            "checkbox": {"default": 1,
                         "command": self.check_methods,
                         "state": "disabled",
                         "cursor": "arrow"
                         }
        }
        self.input_row = EntryRow(methods_frame, 0, **input_row_contents)

        ref_row_contents = {
            "label": {"text": "Reference File",
                      },
            "entry": {"width": 50,
                      "name": "ref_entry",
                      "right_click": self.right_click,
                      "IPTranslator": self.IPTranslator,
                      "callback": self.IPTranslator.check_ref_file
                      },
            "button": {"text": "Browse",
                       "name": "ref_button",
                       "command": browse_file,
                       "args": (("xlsx", "csv", "xls", "xlsm", "xlsb", "odf", "ods", "odt"),
                                "Select a Reference File")
                       },
            "checkbox": {"default": 0,
                         "command": self.check_methods,
                         }
        }
        self.ref_row = EntryRow(methods_frame, 1, **ref_row_contents)

        pan_row_contents = {
            "label": {"text": "PAN Status",
                      },
            "status": {"text": "Not Connected",
                       "foreground": "red",
                       },
            "button": {"text": "Connect",
                       "command": self.get_pan_credentials,
                       },
            "checkbox": {"default": 0,
                         "command": self.check_methods,
                         }
        }
        self.pan_row = StatusRow(methods_frame, 2, **pan_row_contents)

        forti_row_contents = {
            "label": {"text": "Forti Status",
                      },
            "status": {"text": "Not Connected",
                       "foreground": "red",
                       },
            "button": {"text": "Connect",
                       "command": self.get_forti_credentials,
                       },
            "checkbox": {"default": 0,
                         "command": self.check_methods,
                         }
        }
        self.forti_row = StatusRow(methods_frame, 3, **forti_row_contents)

        apic_row_contents = {
            "label": {"text": "APIC Status",
                      },
            "status": {"text": "Not Connected",
                       "foreground": "red",
                       },
            "button": {"text": "Connect",
                       "command": self.get_apic_credentials,
                       },
            "checkbox": {"default": 0,
                         "command": self.check_methods,
                         }
        }
        self.apic_row = StatusRow(methods_frame, 4, **apic_row_contents)

        dns_row_contents = {
            "label": {"text": "DNS Status",
                      },
            "status": {"text": "Not Reachable",
                       "foreground": "red",
                       },
            "button": {"text": "Check",
                       "command": self.get_dns_servers,
                       },
            "checkbox": {"default": 0,
                         "command": self.check_methods,
                         }
        }
        self.dns_row = StatusRow(methods_frame, 5, **dns_row_contents)

        separator(self.__root, orient="horizontal", fill="x", y_offset=(0, 10), x_offset=20)

    def init_root_progressbar(self) -> None:
        """Initialize the progress bar of the application."""

        progress_frame = ttk.Frame(self.__root)
        progress_frame.pack(pady=(0, 10), padx=20, anchor="center")

        self.progress_bar = ttk.Progressbar(progress_frame, orient="horizontal", length=450, mode="determinate")
        self.progress_bar.grid(row=0, column=0, sticky="w")

        self.progress_percentage = ttk.Label(progress_frame, text="%0", font=("Arial", 10, "bold"))
        self.progress_percentage.grid(row=0, column=1, padx=(10, 0), sticky="w")

        separator(self.__root, orient="horizontal", fill="x", y_offset=(0, 10), x_offset=20)

    def init_root_buttons(self) -> None:
        """Initialize the buttons of the application."""

        s = ttk.Style()
        s.configure("start.TButton", font=("Arial", 15, "bold"), foreground="green")

        self.start_button = ttk.Button(self.__root, text="Start", command=self.start,
                                       style="start.TButton",
                                       cursor="hand2")
        self.start_button.pack(pady=(0, 10), anchor="center", ipady=10, ipadx=75)

    def init_root_window(self) -> None:
        """Initialize the root window of the application."""

        self.init_root_header()
        self.init_root_info()
        self.init_root_rows()
        self.init_root_progressbar()
        self.init_root_buttons()

    #######################################################

    def destroy(self, event=None) -> None:
        """Destroy the root window of the application."""
        self.post_start()
        self.IPTranslator.disconnect_pan(False)
        self.IPTranslator.disconnect_forti(False)
        self.IPTranslator.disconnect_apic(False)
        self.__root.destroy()

    def mainloop(self) -> None:
        """Run the main loop of the application."""

        x: int = (self.__SCREEN_WIDTH - self.__width) // 2 - self.__kwargs.get("x_offset", 0)
        y: int = (self.__SCREEN_HEIGHT - self.__height) // 2 - self.__kwargs.get("y_offset", self.__height // 3)

        self.__root.geometry(f"{self.__width}x{self.__height}+{x}+{y}")
        self.__root.resizable(*self.__resizable)

        if self.__icon:
            self.__root.iconphoto(False, self.__icon)

        self.__root.focus_force()
        self.__root.grab_set()

        self.__root.bind('<Escape>', self.destroy)
        self.__root.protocol("WM_DELETE_WINDOW", self.destroy)

        __imported = self.IPTranslator.import_credentials()
        if __imported[0]:
            PropagatingThread(target=self.IPTranslator.connect_pan, daemon=True).start()
        if __imported[1]:
            PropagatingThread(target=self.IPTranslator.connect_forti, daemon=True).start()
        if __imported[2]:
            PropagatingThread(target=self.IPTranslator.connect_apic, daemon=True).start()
        PropagatingThread(target=self.IPTranslator.check_dns_servers, daemon=True).start()

        self.check_methods()

        self.set_info_status("INFO", "green", "arrow")
        self.set_info_message("Welcome to the IP Address Translator App!", "green", "arrow")

        def OnFocusIn(event):
            if type(event.widget).__name__ == 'Tk':
                event.widget.attributes('-topmost', False)
            self.__root.unbind('<FocusIn>', focus_in)

        self.__root.attributes('-topmost', True)
        self.__root.focus_force()
        focus_in = self.__root.bind('<FocusIn>', OnFocusIn)

        self.__root.mainloop()
        self.log.set("INFO - Application closed.")
        self.log.set("")
