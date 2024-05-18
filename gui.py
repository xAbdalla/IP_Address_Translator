import re
import time
import webbrowser
import dns.resolver
from base64 import b64decode
from datetime import datetime
from typing_extensions import Literal

import tkinter as tk
from tkinter import ttk
from tkinter import StringVar
from tkinter import messagebox
from tkinter.filedialog import askopenfilename

import utilities
from utilities import IPTranslator
from utilities import PropagatingThread
from utilities import VerticalScrolledFrame


def browse_file(string_var: StringVar,
                files_ext: list | tuple | set = ("*",),
                title: str = "Select a file") -> str | None:
    files_ext: str = ";".join([f"*.{x}" for x in set(files_ext) if x])

    # Open a file dialog
    filename = askopenfilename(title=title, filetypes=(("Supported Files", files_ext),))

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


def create_rmenu(window=None) -> None | tk.Menu:
    if not window:
        return

    # Create a right click functions
    def cut():
        widget = window.focus_get()
        if isinstance(widget, tk.Entry):
            widget.event_generate("<<Cut>>")

    def copy():
        widget = window.focus_get()
        if isinstance(widget, tk.Entry):
            widget.event_generate("<<Copy>>")

    def paste():
        widget = window.focus_get()
        if isinstance(widget, tk.Entry):
            widget.event_generate("<<Paste>>")

    def delete():
        widget = window.focus_get()
        if isinstance(widget, tk.Entry):
            widget.event_generate("<Delete>")

    def select_all():
        widget = window.focus_get()
        if isinstance(widget, tk.Entry):
            widget.event_generate("<<SelectAll>>")

    def clear_all():
        widget = window.focus_get()
        if isinstance(widget, tk.Entry):
            widget.delete(0, "end")

    # Create a right click menu
    menu = tk.Menu(window, tearoff=0)
    menu.add_command(label="Cut", command=cut)
    menu.add_command(label="Copy", command=copy)
    menu.add_command(label="Paste", command=paste)
    menu.add_command(label="Delete", command=delete)
    menu.add_separator()
    menu.add_command(label="Select All", command=select_all)
    menu.add_command(label="Clear All", command=clear_all)
    return menu


class EntryRow:
    def __init__(self, parent: ttk.Frame, row: int, **kwargs):
        """Create a row with label, entry, button, and checkbox."""

        # Label Variables
        self.label_text = kwargs["label"]["text"]

        # Entry Variables
        self.entry_name = kwargs["entry"]["name"]
        self.entry_width = kwargs["entry"]["width"]
        self.entry_right_click = kwargs["entry"]["right_click"]
        self.entry_callback_fn = kwargs["entry"].get("callback")
        self.translator = kwargs["entry"]["IPTranslator"]

        # Button Variables
        self.button_name = kwargs["button"]["name"]
        self.button_text = kwargs["checkbox"].get("text", "Browse")
        self.button_command = kwargs["button"]["command"]
        self.button_args = kwargs["button"]["args"]

        # Checkbox Variables
        self.checkbox_default = kwargs["checkbox"].get("default", 0)
        self.checkbox_command = kwargs["checkbox"]["command"]
        self.checkbox_state = kwargs["checkbox"].get("state")
        self.checkbox_cursor = kwargs["checkbox"].get("cursor")

        # Create a Label
        self.label = ttk.Label(parent, text=self.label_text)
        self.label.grid(row=row, column=0, sticky="e", padx=(0, 5), pady=(0, 5))

        self.string_var = StringVar()
        if self.entry_callback_fn:
            self.trace_id = self.string_var.trace_add("write", self.entry_callback)

        # Create an Entry
        self.entry = ttk.Entry(parent,
                               textvariable=self.string_var,
                               width=self.entry_width,
                               name=self.entry_name)
        self.entry.bind("<Button-3>", self.entry_right_click)
        if self.entry_callback_fn:
            self.entry.bind("<FocusOut>", self.entry_callback)
        self.entry.grid(row=row, column=1, sticky="w", padx=(0, 5), pady=(0, 5))

        # Create a Button
        self.button = ttk.Button(parent,
                                 text=self.button_text,
                                 command=self.button_callback,
                                 cursor="hand2",
                                 name=self.button_name)
        self.button.grid(row=row, column=2, sticky="w", padx=(0, 5), pady=(0, 5))

        # Create a Combobox
        self.checkbox_var = tk.IntVar()
        self.checkbox = ttk.Checkbutton(parent,
                                        variable=self.checkbox_var,
                                        onvalue=1, offvalue=0,
                                        takefocus=False)
        self.checkbox.var = self.checkbox_var
        self.checkbox_var.set(self.checkbox_default)
        self.checkbox_var.trace_add("write", self.checkbox_command)
        self.checkbox.grid(row=row, column=3, sticky="w", padx=(0, 5), pady=(0, 5))
        if self.checkbox_state and self.checkbox_cursor:
            self.checkbox.config(state=self.checkbox_state, cursor=self.checkbox_cursor)

        self.thread = PropagatingThread(target=None, daemon=True)

    def entry_callback(self, *args) -> None:
        self.thread = PropagatingThread(target=self.entry_callback_fn, args=(str(self.entry.focus_get()),), daemon=True)
        self.thread.start()

    def button_callback(self) -> str | None:
        filename = self.button_command(self.string_var, *self.button_args)
        if filename:
            if self.entry_name == "input_entry":
                self.translator.input_file = filename
            elif self.entry_name == "ref_entry":
                self.translator.ref_file = filename
        return filename


class StatusRow:
    def __init__(self, parent: ttk.Frame, row: int, **kwargs):
        """Create a row with label, status_label, button, and checkbox."""

        # Label Variables
        self.label_text = kwargs["label"]["text"]

        # Status Variables
        self.status_text = kwargs["status"]["text"]
        self.status_foreground = kwargs["status"]["foreground"]

        # Button Variables
        self.button_text = kwargs["button"]["text"]
        self.button_command = kwargs["button"]["command"]

        # Checkbox Variables
        self.checkbox_default = kwargs["checkbox"].get("default", 0)
        self.checkbox_command = kwargs["checkbox"]["command"]
        self.checkbox_state = kwargs["checkbox"].get("state")
        self.checkbox_cursor = kwargs["checkbox"].get("cursor")

        # Create a Label
        self.label = ttk.Label(parent, text=self.label_text)
        self.label.grid(row=row, column=0, sticky="e", padx=(0, 5), pady=(0, 5))

        # Create a status Label
        self.status = ttk.Label(parent,
                                text=self.status_text,
                                foreground=self.status_foreground,
                                font=("Arial", 10, "bold"))
        self.status.grid(row=row, column=1, sticky="w", padx=(0, 5), pady=(0, 5))

        self.button = ttk.Button(parent,
                                 text=self.button_text,
                                 command=self.button_command,
                                 cursor="hand2")
        self.button.grid(row=row, column=2, sticky="w", padx=(0, 5), pady=(0, 5))

        # Create a Combobox
        self.checkbox_var = tk.IntVar()
        self.checkbox = ttk.Checkbutton(parent,
                                        variable=self.checkbox_var,
                                        onvalue=1, offvalue=0,
                                        takefocus=False,
                                        )
        self.checkbox.var = self.checkbox_var
        self.checkbox_var.set(self.checkbox_default)
        self.checkbox_var.trace_add("write", self.checkbox_command)
        self.checkbox.grid(row=row, column=3, sticky="w", padx=(0, 5), pady=(0, 5))
        if self.checkbox_state and self.checkbox_cursor:
            self.checkbox.config(state=self.checkbox_state, cursor=self.checkbox_cursor)


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

        self.parent = parent
        self.title = title
        self.icon = icon

        self.app = app
        self.rows = rows
        self.translator = translator

        self.window = tk.Toplevel(self.parent)
        self.window.title(self.title)

        self.right_click_menu = create_rmenu(self.window)

        self.window.focus_force()
        self.window.grab_set()
        self.window.focus_set()
        self.window.transient(self.parent)

        WIDTH = kwargs.get("width", 350)
        HEIGHT = kwargs.get("height", 230)

        self.window.geometry(f'{WIDTH}x{HEIGHT}+{self.parent.winfo_rootx() + 100}+{self.parent.winfo_rooty() + 75}')
        self.window.resizable(False, False)

        self.__frame = ttk.Frame(self.window)
        self.__frame.pack(fill="both", expand=True)

        self.title_label = ttk.Label(self.__frame, text=self.title, font=("Arial", 10, "bold"))
        self.title_label.grid(row=0, column=0, columnspan=2, padx=5, pady=5)

        # row = ["text", "default", "bold"]
        self.elements = {}
        self.entries_var = [StringVar() for _ in range(len(self.rows))]
        for i, row in enumerate(self.rows):
            self.entries_var.append(StringVar())
            name, default, bold = row

            self.elements[name] = {}
            self.elements[name]["label"] = ttk.Label(self.__frame, text=name,
                                                     font=("Arial", 8, "bold" if bold else "normal"))
            self.elements[name]["label"].grid(row=i + 1, column=0, sticky="e", padx=5, pady=5)
            self.elements[name]["entry"] = ttk.Entry(self.__frame, textvariable=self.entries_var[i], width=37)
            self.elements[name]["entry"].bind("<Button-3>", self.right_click)
            if "password" in name.lower():
                self.elements[name]["entry"].config(show="*")
            self.entries_var[i].set(default)
            self.elements[name]["entry"].grid(row=i + 1, column=1, sticky="w", padx=5, pady=5)

        self.remember_me_var = tk.IntVar()
        self.remember_checkbox = ttk.Checkbutton(self.__frame,
                                                 text="Remember Me",
                                                 variable=self.remember_me_var,
                                                 onvalue=1, offvalue=0,
                                                 cursor="hand2",
                                                 takefocus=False,
                                                 width=15)
        self.remember_checkbox.grid(row=len(self.rows) + 1, column=0, columnspan=2, sticky="w", padx=25, pady=5)

        self.__buttons_frame = ttk.Frame(self.__frame)
        self.__buttons_frame.grid(row=len(self.rows) + 2, column=0, columnspan=2, padx=5, pady=5)

        self.save_button = ttk.Button(self.__buttons_frame,
                                      text="Save",
                                      command=self.save,
                                      cursor="hand2",
                                      width=15)
        self.save_button.grid(row=0, column=0, sticky="w", padx=(20, 5), pady=5)

        self.cancel_button = ttk.Button(self.__buttons_frame,
                                        text="Cancel",
                                        command=self.destroy,
                                        cursor="hand2",
                                        width=15)
        self.cancel_button.grid(row=0, column=1, sticky="e", padx=5, pady=5)

    def right_click(self, event) -> None:
        event.widget.focus()
        self.right_click_menu.tk_popup(event.x_root, event.y_root)

    def save(self) -> bool | None:
        self.save_button.config(state="disabled", cursor="arrow")
        self.cancel_button.config(state="disabled", cursor="arrow")
        self.window.protocol("WM_DELETE_WINDOW", lambda: False)

        __strings = []
        remember = self.remember_me_var.get()
        for value in self.entries_var:
            __strings.append(value.get().strip())

        if self.app in ["pan", "forti", "apic"]:
            __connected = None
            if self.app in ["pan", "apic"]:
                __ip = __strings[0]
                __username = __strings[1]
                __password = __strings[2]
                __vsys = __strings[3]

                if self.app == "apic" and not __vsys:
                    messagebox.showerror("Error",
                                         "Class(es) can not be empty.",
                                         parent=self.window)
                    self.save_button.config(state="normal", cursor="hand2")
                    self.cancel_button.config(state="normal", cursor="hand2")
                    self.window.protocol("WM_DELETE_WINDOW", self.destroy)
                    return False

                if self.app == "pan":
                    __connected = PropagatingThread(target=self.translator.connect_pan,
                                                    kwargs={"ip": __ip,
                                                            "username": __username,
                                                            "password": __password,
                                                            "vsys": __vsys,
                                                            "window": self.window},
                                                    daemon=True,
                                                    name="connect_pan")
                elif self.app == "apic":
                    __connected = PropagatingThread(target=self.translator.connect_apic,
                                                    kwargs={"ip": __ip,
                                                            "username": __username,
                                                            "password": __password,
                                                            "apic_class": __vsys,
                                                            "window": self.window},
                                                    daemon=True,
                                                    name="connect_apic")

            elif self.app == "forti":
                __ip = __strings[0]
                __port = __strings[1]
                __username = __strings[2]
                __password = __strings[3]
                __vdom = __strings[4]

                __connected = PropagatingThread(target=self.translator.connect_forti,
                                                kwargs={"ip": __ip,
                                                        "port": __port,
                                                        "username": __username,
                                                        "password": __password,
                                                        "vdom": __vdom,
                                                        "window": self.window},
                                                daemon=True,
                                                name="connect_forti")

            __connected.start()
            while __connected.is_alive():
                self.window.update()
                self.parent.update()
                time.sleep(0.1)

            __connected = __connected.join()

            if not __connected:
                if not tk.messagebox.askyesno("Warning",
                                              "Connection failed, do you want to save the credentials anyway?",
                                              icon="warning",
                                              default="no",
                                              parent=self.window):
                    self.save_button.config(state="normal", cursor="hand2")
                    self.cancel_button.config(state="normal", cursor="hand2")
                    self.window.protocol("WM_DELETE_WINDOW", self.destroy)
                    return False

            if self.app == "pan":
                self.translator.pan_ip = __ip
                self.translator.pan_username = __username
                self.translator.pan_password = __password
                self.translator.pan_vsys = __vsys
                if not __connected:
                    self.translator.pan_addresses = []

                if remember:
                    self.translator.save_credentials(app="pan")

            elif self.app == "forti":
                self.translator.forti_ip = __ip
                self.translator.forti_port = __port
                self.translator.forti_username = __username
                self.translator.forti_password = __password
                self.translator.forti_vdom = __vdom
                if not __connected:
                    self.translator.forti_addresses = []

                if remember:
                    self.translator.save_credentials(app="forti")

            elif self.app == "apic":
                self.translator.apic_ip = __ip
                self.translator.apic_username = __username
                self.translator.apic_password = __password
                self.translator.apic_class = __vsys
                if not __connected:
                    self.translator.apic_addresses = []

                if remember:
                    self.translator.save_credentials(app="apic")

        elif self.app == "dns":
            __servers = __strings.copy()

            for i in range(3, -1, -1):
                if not __servers[i]:
                    __servers.pop(i)
            if len(__servers) < 4:
                __servers.extend([""] * (4 - len(__servers)))

            if not __servers:
                messagebox.showerror("Error",
                                     "DNS servers can not be empty.",
                                     parent=self.window)
                self.save_button.config(state="normal", cursor="hand2")
                self.cancel_button.config(state="normal", cursor="hand2")
                self.window.protocol("WM_DELETE_WINDOW", self.destroy)
                return False

            __connected = PropagatingThread(target=self.translator.check_dns_servers,
                                            args=(__servers,),
                                            daemon=True,
                                            name="check_dns_servers")
            __connected.start()
            while __connected.is_alive():
                self.window.update()
                self.parent.update()
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
                                           parent=self.window):
                    self.save_button.config(state="normal", cursor="hand2")
                    self.cancel_button.config(state="normal", cursor="hand2")
                    self.window.protocol("WM_DELETE_WINDOW", self.destroy)
                    return False

            elif len(__bad_servers):
                if not messagebox.askyesno("Warning",
                                           f"Invalid DNS servers found:\n    {'\n    '.join(__bad_servers)}\n\nDo you want to save the servers anyway?",
                                           icon="warning",
                                           parent=self.window):
                    self.save_button.config(state="normal", cursor="hand2")
                    self.cancel_button.config(state="normal", cursor="hand2")
                    self.window.protocol("WM_DELETE_WINDOW", self.destroy)
                    return False

            self.translator.dns_servers = __servers.copy()
            self.translator.resolvers = __good_servers.copy()

            if remember:
                self.translator.save_credentials(app="dns")

        self.destroy()

    def destroy(self) -> None:
        self.window.destroy()
        self.parent.attributes('-disabled', False)
        self.parent.update()
        self.parent.focus_force()

    def mainloop(self) -> None:
        self.parent.attributes('-disabled', True)
        self.window.protocol("WM_DELETE_WINDOW", self.destroy)
        self.window.bind('<Escape>', self.destroy)
        self.window.bind('<Return>', self.save)
        self.window.iconphoto(False, self.icon)
        self.window.mainloop()


class GUI:
    """A class that represents the windows gui of the application."""

    def __init__(self,
                 name: str = "IP Address Translator",
                 short_name: str = "IAT",
                 version: str = "1.5",
                 width: int = 550,
                 height: int = 490,
                 icon_data: str = utilities.ICON_DATA,
                 *args, **kwargs):
        """Initialize the GUI class."""
        self.DEBUG = False

        self.name: str = name
        self.shorten_name: str = short_name
        self.version: str = version

        self.__width: int = width
        self.__height: int = height
        self.__resizable: tuple[bool, bool] = (False, False)
        self.__args = args
        self.__kwargs = kwargs

        self.author: str = "xAbdalla"
        self.github_link: str = "https://github.com/xAbdalla"
        self.github_repo_link: str = f"{self.github_link}/{self.name.replace(" ", "_")}"

        self.full_name: str = f"{self.name} ({self.shorten_name})"
        self.title: str = f"{self.full_name} v{self.version} by {self.author}"
        self.title_label: str = f"{self.name} v{self.version}"

        self.root: tk.Tk = tk.Tk()
        self.root.title(self.title)

        self.icon: tk.PhotoImage = tk.PhotoImage(data=b64decode(icon_data))
        self.github_icon: tk.PhotoImage = tk.PhotoImage(data=b64decode(utilities.GITHUB_ICON_DATA))
        self.github_icon = self.github_icon.zoom(1).subsample(18)

        self.info_icon: tk.PhotoImage = tk.PhotoImage(data=b64decode(utilities.INFO_ICON_DATA))
        self.info_icon = self.info_icon.zoom(2).subsample(13)

        self.__SCREEN_WIDTH = self.root.winfo_screenwidth()
        self.__SCREEN_HEIGHT = self.root.winfo_screenheight()

        self.right_click_menu = create_rmenu(self.root)

        self.log = StringVar()
        if self.__kwargs.get("log"):
            self.log.trace_add("write", self.write_logs)

            # Check if the log file exists and big delete it
            # if os.path.exists("logs.txt"):
            #     if os.path.getsize("logs.txt") > int(5 * 1024 * 1024):
            #         os.remove("logs.txt")
        self.log.set("INFO - Application started.")

        self.IPTranslator: IPTranslator = IPTranslator(parent=self, debug=self.DEBUG)
        self.start_thread: PropagatingThread | None = None
        self.start_flag: bool = False
        self.pause_flag: bool = False
        self.stop_flag: bool = False
        self.methods_flags: list[bool] = [False for _ in range(5)]

        self.info_type_label = None
        self.msg_text = StringVar()
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

    def right_click(self, event) -> None:
        event.widget.focus()
        self.right_click_menu.tk_popup(event.x_root, event.y_root)

    def write_logs(self, *arg):
        log = self.log.get()
        with open("logs.txt", "a") as f:
            if log:
                f.write(f"{datetime.now().strftime("%Y-%m-%d %H:%M:%S")} - {log}\n")
            else:
                f.write(f"{"=" * 50}\n")

    def disable_buttons(self, button: str = "") -> None:
        if button == "all":
            self.root.protocol("WM_DELETE_WINDOW", self.stop)
            button = ""

        if button in ["input", ""]:
            self.input_row.button.config(state="disabled", cursor="arrow")
            self.input_row.entry.config(state="disabled", cursor="arrow")
            self.input_row.entry.unbind("<FocusOut>")
        if button in ["ref", ""]:
            self.ref_row.button.config(state="disabled", cursor="arrow")
            self.ref_row.entry.config(state="disabled", cursor="arrow")
            self.ref_row.checkbox.config(state="disabled", cursor="arrow")
            self.ref_row.entry.unbind("<FocusOut>")
        if button in ["pan", ""]:
            self.pan_row.button.config(state="disabled", cursor="arrow")
            self.pan_row.checkbox.config(state="disabled", cursor="arrow")
        if button in ["forti", ""]:
            self.forti_row.button.config(state="disabled", cursor="arrow")
            self.forti_row.checkbox.config(state="disabled", cursor="arrow")
        if button in ["apic", ""]:
            self.apic_row.button.config(state="disabled", cursor="arrow")
            self.apic_row.checkbox.config(state="disabled", cursor="arrow")
        if button in ["dns", ""]:
            self.dns_row.button.config(state="disabled", cursor="arrow")
            self.dns_row.checkbox.config(state="disabled", cursor="arrow")

        if not self.start_flag:
            self.start_button.config(state="disabled", cursor="arrow")
            self.root.unbind("<Return>")

    def enable_buttons(self, button: str = "") -> None:
        if button == "all":
            self.root.protocol("WM_DELETE_WINDOW", self.destroy)
            button = ""

        if button in ["input", ""]:
            if button == "input" or (self.input_row.thread and not self.input_row.thread.is_alive()):
                self.input_row.button.config(state="normal", cursor="hand2")
                self.input_row.entry.config(state="normal", cursor="xterm")
                self.input_row.entry.bind("<FocusOut>", self.input_row.entry_callback)
        if button in ["ref", ""]:
            if button == "ref" or (self.ref_row.thread and not self.ref_row.thread.is_alive()):
                self.ref_row.button.config(state="normal", cursor="hand2")
                self.ref_row.entry.config(state="normal", cursor="xterm")
                self.ref_row.checkbox.config(state="normal", cursor="hand2")
                self.ref_row.entry.bind("<FocusOut>", self.ref_row.entry_callback)
        if button in ["pan", ""]:
            if self.pan_row.button["text"] == "Connect":
                self.pan_row.button.config(state="normal", cursor="hand2")
                self.pan_row.checkbox.config(state="normal", cursor="hand2")
        if button in ["forti", ""]:
            if self.forti_row.button["text"] == "Connect":
                self.forti_row.button.config(state="normal", cursor="hand2")
                self.forti_row.checkbox.config(state="normal", cursor="hand2")
        if button in ["apic", ""]:
            if self.apic_row.button["text"] == "Connect":
                self.apic_row.button.config(state="normal", cursor="hand2")
                self.apic_row.checkbox.config(state="normal", cursor="hand2")
        if button in ["dns", ""]:
            if self.dns_row.button["text"] == "Check":
                self.dns_row.button.config(state="normal", cursor="hand2")
                self.dns_row.checkbox.config(state="normal", cursor="hand2")

        __input = str(self.input_row.button["state"]) == "normal"
        __ref = str(self.ref_row.button["state"]) == "normal"
        __pan = str(self.pan_row.button["state"]) == "normal"
        __forti = str(self.forti_row.button["state"]) == "normal"
        __apic = str(self.apic_row.button["state"]) == "normal"
        __dns = str(self.dns_row.button["state"]) == "normal"

        if all([__input, __ref, __pan, __forti, __apic, __dns, True in self.methods_flags]):
            self.start_button.config(state="normal", cursor="hand2")
            self.root.bind("<Return>", self.start)

    def set_info_status(self, status: str, foreground: str, cursor: str) -> None:
        """Set the status of the info label."""
        self.info_type_label.config(text=status, foreground=foreground, cursor=cursor)

    def set_info_message(self, message: str, foreground: str, cursor: str) -> None:
        """Set the message of the info label."""
        self.msg_text.set(message)
        self.info_msg_label.config(foreground=foreground, cursor=cursor)

    def check_methods(self, event=None, *args) -> bool:
        if self.ref_row.checkbox_var.get():
            self.methods_flags[0] = True
        else:
            self.methods_flags[0] = False

        if self.pan_row.checkbox_var.get():
            self.methods_flags[1] = True
        else:
            self.methods_flags[1] = False

        if self.forti_row.checkbox_var.get():
            self.methods_flags[2] = True
        else:
            self.methods_flags[2] = False

        if self.apic_row.checkbox_var.get():
            self.methods_flags[3] = True
        else:
            self.methods_flags[3] = False

        if self.dns_row.checkbox_var.get():
            self.methods_flags[4] = True
        else:
            self.methods_flags[4] = False

        if True in self.methods_flags:
            if not self.start_flag:
                self.set_info_status("INFO", "green", "arrow")
                self.set_info_message("Click the (Start) button to start the translation process",
                                      "green", "arrow")

            self.enable_buttons("start")
            return True

        if not self.start_flag:
            self.set_info_status("INFO", "green", "arrow")
            self.set_info_message("Please select one searching method at least",
                                  "red", "arrow")

        self.disable_buttons("start")
        return False

    #######################################################

    def info_window(self) -> None:
        def destroy_info_window(event=None):
            info_window.destroy()
            self.root.attributes('-disabled', False)
            self.root.update()
            self.root.focus_force()

        font_bold = ("Arial", 10, "bold")
        font_normal = ("Arial", 10)

        objective_text = f"""{self.name} application assists you in mapping IPs from network logs to descriptive \
object names or domain names, making log analysis more straightforward."""
        features_text = f"""➞ Various Input Options:
      • Accepts Excel, CSV, or Text files.
      • Direct input of a single IP address, subnet, range, or list of them separated by comma.

➞ Files Specifications:
      • All files must have its proper extension.
      • Input Text files must have one subnet per line.
      • Input CSV files must have a column named 'Subnet'.
      • Input Excel files could have multiple sheets; each must contain a column named 'Subnet'.
      • Reference CSV files require three columns: 'Tenant', 'Address object', and 'Subnet'.
      • Reference Excel files could have multiple sheets, each must contain the three columns.

➞ Various Searching Methods:
      • Reference File: Searches for matches in a user-provided reference file.
      • Palo Alto: Connects via Palo Alto device REST API to import address objects.
      • Fortinet: Connects via FortiGate REST API to retrieve address objects.
      • Cisco ACI: Connects via SSH to APIC and import address objects based on the specified Class.
      • Reverse DNS: Resolves IPs to domain names using the PTR records.

➞ Palo Alto Device Specifications:
      • Ensure Panorama/Firewall is reachable and you have REST API access/enable.
      • Leave "VSYS" field empty if you want to import address objects from all virtual systems.

➞ Fortinet FortiGate Specifications:
      • Ensure FortiGate is reachable and you have REST API access/enable.
      • Leave "VDOM" field empty if you want to import address objects from all virtual domains.

➞ Cisco ACI Specifications:
      • Ensure APIC is reachable and has CLI access.
      • Specify the Class(es) of the address objects to be searched.
      • The program searches the "dn" attribute exclusively.

➞ DNS Resolver:
      • Resolves IPs to domain names using the system DNS servers or a user-provided DNS servers.
      • You can specify up to four DNS servers.

➞ Invalid objects:
      • FQDN Objects.
      • Object name is a valid IP or subnet.
      • Object name is "network_" and a valid IP or subnet.
      • Object address is 0.0.0.0/0 or 0.0.0.0/32.

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

        info_window = tk.Toplevel(self.root)
        info_window.title("Help")
        info_window.focus_force()
        info_window.grab_set()
        info_window.focus_set()
        info_window.transient(self.root)
        self.root.attributes('-disabled', True)

        info_window.geometry(f'{665}x{580}+{self.root.winfo_rootx() - 65}+{self.root.winfo_rooty() - 100}')
        info_window.resizable(False, False)
        info_window.iconphoto(False, self.info_icon)

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

        author_label = ttk.Label(info_frame.interior, text=f"{self.title}", font=font_bold)
        author_label.pack(pady=(10, 0), padx=10, anchor="w")

        feedback_label = ttk.Label(info_frame.interior, text=f"Feel free to provide feedback or report issues.",
                                   font=font_normal)
        feedback_label.pack(pady=0, padx=10, anchor="w")

        github_label = ttk.Label(info_frame.interior, text=f"GitHub",
                                 font=font_bold, cursor="hand2", foreground="blue")
        github_label.pack(pady=(0, 5), padx=10, anchor="w")
        github_label.bind("<Button-1>", lambda e: webbrowser.open_new(self.github_repo_link))

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

        __pan_window = CredentialsWindow(parent=self.root,
                                         title="PAN Credentials",
                                         app="pan",
                                         rows=__rows,
                                         translator=self.IPTranslator,
                                         icon=self.icon,
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

        __forti_window = CredentialsWindow(parent=self.root,
                                           title="Forti Credentials",
                                           app="forti",
                                           rows=__rows,
                                           translator=self.IPTranslator,
                                           icon=self.icon,
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

        __apic_window = CredentialsWindow(parent=self.root,
                                          title="APIC Credentials",
                                          app="apic",
                                          rows=__rows,
                                          translator=self.IPTranslator,
                                          icon=self.icon,
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

        __dns_window = CredentialsWindow(parent=self.root,
                                         title="DNS Servers",
                                         app="dns",
                                         rows=__rows,
                                         translator=self.IPTranslator,
                                         icon=self.icon,
                                         width=305,
                                         height=230
                                         )
        __dns_window.mainloop()

    #######################################################

    def pre_start(self) -> None:
        self.post_start()
        self.start_flag = True
        self.pause_flag = False
        self.stop_flag = False

        s = ttk.Style()
        s.configure("start.TButton", font=("Arial bold", 15), foreground="red")

        self.start_button.config(text="Stop", state="normal", style="start.TButton", command=self.stop, cursor="hand2")
        self.root.bind("<Return>", self.stop)

        self.disable_buttons("all")

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
                self.root.update()
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
            self.set_info_message(f"Input is valid. {self.IPTranslator.all_inputs_no} address(es) found.",
                                  "green", "arrow")
            self.log.set(f"INFO - Input is valid. {self.IPTranslator.all_inputs_no} address(es) found.")

        # Check Used Methods
        self.check_methods()

        if not (True in self.methods_flags):
            self.post_start()
            self.set_info_status("ERROR", "red", "arrow")
            self.set_info_message("Please select one searching method at least", "red", "arrow")
            return False

        if self.methods_flags[0]:
            # Check Reference File
            if not self.IPTranslator.refs:
                __ref_thread = PropagatingThread(target=self.IPTranslator.check_ref_file,
                                                 daemon=True)
                __ref_thread.start()
                while __ref_thread.is_alive():
                    self.root.update()
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

        if self.methods_flags[1]:
            if not self.IPTranslator.pan_addresses:
                # Connect to Palo Alto
                __connect_pan_thread = PropagatingThread(target=self.IPTranslator.connect_pan,
                                                         daemon=True)
                __connect_pan_thread.start()
                while __connect_pan_thread.is_alive():
                    self.root.update()
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
                        self.root.update()
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

        if self.methods_flags[2]:
            if not self.IPTranslator.forti_addresses:
                # Connect to Fortinet
                __connect_forti_thread = PropagatingThread(target=self.IPTranslator.connect_forti,
                                                           daemon=True)
                __connect_forti_thread.start()
                while __connect_forti_thread.is_alive():
                    self.root.update()
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
                        self.root.update()
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

        if self.methods_flags[3]:
            if not self.IPTranslator.apic_addresses:
                # Connect to Cisco ACI
                __connect_apic_thread = PropagatingThread(target=self.IPTranslator.connect_apic,
                                                          daemon=True)
                __connect_apic_thread.start()
                while __connect_apic_thread.is_alive():
                    self.root.update()
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
                        self.root.update()
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

        if self.methods_flags[4]:
            if not self.IPTranslator.resolvers:
                # Get DNS Servers
                __dns_thread = PropagatingThread(target=self.IPTranslator.check_dns_servers,
                                                 daemon=True)
                __dns_thread.start()
                while __dns_thread.is_alive():
                    self.root.update()
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

        if self.methods_flags[0]:
            if len(self.IPTranslator.refs.keys()) == 1:
                __sheet_str = "1 Sheet"
            else:
                __sheet_str = f"{len(self.IPTranslator.refs.keys())} Sheets"

            __content_str += f"Reference File Contents:\n{__sheet_str}, {self.IPTranslator.ref_no} Valid Subnets, {len(self.IPTranslator.ref_invalids)} Invalid Subnets."
            if self.IPTranslator.ref_invalids:
                __content_str += f"\nInvalid details are saved in the 'invalid_records.txt' file.\n\n"
            else:
                __content_str += "\n\n"

        if self.methods_flags[1]:
            __content_str += f"Palo Alto: {len(self.IPTranslator.pan_addresses)} Address Objects.\n\n"

        if self.methods_flags[2]:
            __content_str += f"FortiGate: {len(self.IPTranslator.forti_addresses)} Address Objects.\n\n"

        if self.methods_flags[3]:
            __content_str += f"Cisco ACI: {len(self.IPTranslator.apic_addresses)} Address Objects.\n\n"

        if self.methods_flags[4]:
            __content_str += f"DNS Servers: {len(self.IPTranslator.resolvers)} Servers.\n\n"

        ans = messagebox.askquestion("Info",
                                     f"{__content_str}\nClick 'Yes' to start the translation process.",
                                     parent=self.root)

        if ans != "yes":
            self.set_info_status("INFO", "green", "arrow")
            self.set_info_message("The process has been canceled.", "red", "arrow")
            self.log.set("INFO - The process has been canceled by the user.")

            self.post_start()
            return False

        self.start_thread = PropagatingThread(target=self.IPTranslator.Translate)
        self.start_thread.start()

    def post_start(self) -> None:
        self.start_flag = False
        self.pause_flag = False
        self.stop_flag = False

        try:
            if self.start_thread.is_alive():
                self.start_thread.stop()
        except:
            pass

        s = ttk.Style()
        s.configure("start.TButton", font=("Arial bold", 15), foreground="green")

        self.start_button.config(text="Start", state="normal", style="start.TButton", command=self.start,
                                 cursor="hand2")
        self.root.bind("<Return>", self.start)

        self.enable_buttons("all")

    def stop(self, event=None) -> None:
        self.pause_flag = True

        ans = messagebox.askyesnocancel("Warning",
                                        "Do you want to save the results before stopping the process?",
                                        icon="warning",
                                        default="yes",
                                        parent=self.root)

        if ans is True:
            self.log.set("INFO - User chose to stop the process and save the results.")
            self.start_flag = False

        elif ans is False:
            self.stop_flag = True
            self.log.set("INFO - User chose to stop the process without saving the results.")

            if self.progress_bar['value'] >= self.progress_bar['maximum']:
                self.progress_percentage.config(text="100%", foreground="green")
            else:
                self.progress_percentage.config(foreground="red")

            self.post_start()
            self.IPTranslator.clear_var()
            self.root.update()

            self.set_info_status("INFO", "green", "arrow")
            self.set_info_message("The process has been stopped successfully.", "red", "arrow")

        self.pause_flag = False

    #######################################################

    def init_root_header(self) -> None:
        """Initialize the header of the application."""

        header = ttk.Label(self.root, text=self.name, font=("Arial", 20, "bold"))
        header.pack(pady=(15, 0), anchor="center")

        head_btn_frame = ttk.Frame(self.root)
        head_btn_frame.pack(pady=(0, 10), padx=20, anchor="center")

        github_button = tk.Button(head_btn_frame, text=f" {self.author}", font=("Arial bold", 12),
                                  image=self.github_icon, borderwidth=0, relief="sunken", compound="left",
                                  cursor="hand2", command=lambda: webbrowser.open_new(self.github_link),
                                  anchor="center")
        github_button.grid(row=0, column=0, padx=(5, 300), pady=(3, 5), sticky="w")

        info_button = tk.Button(head_btn_frame, text="Help ", font=("Arial bold", 12),
                                image=self.info_icon, borderwidth=0, relief="sunken", compound="right",
                                cursor="hand2", command=lambda: self.info_window(), anchor="center")
        info_button.grid(row=0, column=1, padx=(0, 5), pady=(3, 5), sticky="e")

        separator(self.root, orient="horizontal", fill="x", y_offset=(0, 10), x_offset=20)

    def init_root_info(self) -> None:
        """Initialize the info frame of the application."""

        info_frame = ttk.Frame(self.root)
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

        separator(self.root, orient="horizontal", fill="x", y_offset=(0, 10), x_offset=20)

    def init_root_rows(self) -> None:
        """Initialize the rows of the application."""

        methods_frame = ttk.Frame(self.root)
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

        separator(self.root, orient="horizontal", fill="x", y_offset=(0, 10), x_offset=20)

    def init_root_progressbar(self) -> None:
        """Initialize the progress bar of the application."""

        progress_frame = ttk.Frame(self.root)
        progress_frame.pack(pady=(0, 10), padx=20, anchor="center")

        self.progress_bar = ttk.Progressbar(progress_frame, orient="horizontal", length=450, mode="determinate")
        self.progress_bar.grid(row=0, column=0, sticky="w")

        self.progress_percentage = ttk.Label(progress_frame, text="%0", font=("Arial", 10, "bold"))
        self.progress_percentage.grid(row=0, column=1, padx=(10, 0), sticky="w")

        separator(self.root, orient="horizontal", fill="x", y_offset=(0, 10), x_offset=20)

    def init_root_buttons(self) -> None:
        """Initialize the buttons of the application."""

        s = ttk.Style()
        s.configure("start.TButton", font=("Arial", 15, "bold"), foreground="green")

        self.start_button = ttk.Button(self.root, text="Start", command=self.start,
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
        self.root.destroy()

    def mainloop(self) -> None:
        """Run the main loop of the application."""

        x: int = (self.__SCREEN_WIDTH - self.__width) // 2 - self.__kwargs.get("x_offset", 0)
        y: int = (self.__SCREEN_HEIGHT - self.__height) // 2 - self.__kwargs.get("y_offset", self.__height // 3)

        self.root.geometry(f"{self.__width}x{self.__height}+{x}+{y}")
        self.root.resizable(*self.__resizable)

        if self.icon:
            self.root.iconphoto(False, self.icon)

        self.root.focus_force()
        self.root.grab_set()

        self.root.bind('<Escape>', self.destroy)
        self.root.protocol("WM_DELETE_WINDOW", self.destroy)

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
            self.root.unbind('<FocusIn>', focus_in)

        self.root.attributes('-topmost', True)
        self.root.focus_force()
        focus_in = self.root.bind('<FocusIn>', OnFocusIn)

        try:
            self.root.mainloop()
            self.log.set("INFO - Application closed successfully.")
        except:
            self.log.set("INFO - Application did not close properly.")
        self.log.set("")
