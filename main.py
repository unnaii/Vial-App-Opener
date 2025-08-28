import os
import sys
import json
import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import keyboard
import subprocess
import winreg
import pythoncom
from win32com.shell import shell, shellcon

APP_NAME = "VialMacropad"
CONFIG_DIR = os.path.join(os.getenv('APPDATA'), APP_NAME)
CONFIG_FILE = os.path.join(CONFIG_DIR, "config_macropad.json")

HOTKEYS = ['f13', 'f14', 'f15', 'f16', 'f17', 'f18', 'f19', 'f20', 'f21', 'f22', 'f23', 'f24']
ASSIGNABLE_HOTKEYS = [hk for hk in HOTKEYS if hk != 'f24']

def cargar_config():
    if os.path.exists(CONFIG_FILE):
        try:
            with open(CONFIG_FILE, "r") as f:
                return json.load(f)
        except:
            return {}
    return {}

def guardar_config(config):
    if not os.path.exists(CONFIG_DIR):
        os.makedirs(CONFIG_DIR)
    with open(CONFIG_FILE, "w") as f:
        json.dump(config, f, indent=4)

def ejecutar_programa(ruta):
    try:
        subprocess.Popen(ruta)
    except Exception as e:
        print(f"No se pudo ejecutar {ruta}: {e}")

def obtener_programas_menu_inicio():
    pythoncom.CoInitialize()

    carpetas = [
        shell.SHGetFolderPath(0, shellcon.CSIDL_PROGRAMS, None, 0),
        shell.SHGetFolderPath(0, shellcon.CSIDL_COMMON_PROGRAMS, None, 0)
    ]

    accesos = {}
    for carpeta in carpetas:
        for root, dirs, files in os.walk(carpeta):
            for file in files:
                if file.lower().endswith(".lnk"):
                    try:
                        ruta = os.path.join(root, file)
                        nombre = os.path.splitext(file)[0]
                        accesos[nombre] = ruta
                    except Exception:
                        pass

    return accesos

def resolver_acceso_directo(ruta_lnk):
    pythoncom.CoInitialize()
    shell_link = pythoncom.CoCreateInstance(shell.CLSID_ShellLink, None,
                                           pythoncom.CLSCTX_INPROC_SERVER,
                                           shell.IID_IShellLink)
    persist_file = shell_link.QueryInterface(pythoncom.IID_IPersistFile)
    persist_file.Load(ruta_lnk)
    return shell_link.GetPath(shell.SLGP_UNCPRIORITY)[0]

def agregar_al_inicio(nombre_app, ruta_exe):
    clave = r"Software\Microsoft\Windows\CurrentVersion\Run"
    try:
        with winreg.OpenKey(winreg.HKEY_CURRENT_USER, clave, 0, winreg.KEY_ALL_ACCESS) as reg_key:
            winreg.SetValueEx(reg_key, nombre_app, 0, winreg.REG_SZ, ruta_exe)
        print("Agregado al inicio.")
    except Exception as e:
        print("Error al agregar al inicio:", e)

def quitar_del_inicio(nombre_app):
    clave = r"Software\Microsoft\Windows\CurrentVersion\Run"
    try:
        with winreg.OpenKey(winreg.HKEY_CURRENT_USER, clave, 0, winreg.KEY_ALL_ACCESS) as reg_key:
            winreg.DeleteValue(reg_key, nombre_app)
        print("Eliminado del inicio.")
    except FileNotFoundError:
        print("No estaba agregado al inicio.")
    except Exception as e:
        print("Error al eliminar del inicio:", e)

class App:
    def __init__(self, root):
        self.root = root
        self.root.title("Asignador teclas F13-F23 para Macropad (Vial) | F24 para minimizar")
        self.config = cargar_config()
        self.labels = {}
        self.handlers = {}

        for i, hotkey in enumerate(ASSIGNABLE_HOTKEYS):
            tk.Label(root, text=hotkey.upper()).grid(row=i, column=0, padx=5, pady=5, sticky="w")

            ruta_actual = self.config.get(hotkey, "(no asignado)")
            label = tk.Label(root, text=os.path.basename(ruta_actual) if os.path.exists(ruta_actual) else ruta_actual)
            label.grid(row=i, column=1, padx=5, pady=5, sticky="w")
            self.labels[hotkey] = label

            btn = tk.Button(root, text="Asignar programa", command=lambda h=hotkey, l=label: self.asignar_programa_desde_lista(h, l))
            btn.grid(row=i, column=2, padx=5, pady=5)
        
        tk.Label(root, text="F24").grid(row=len(ASSIGNABLE_HOTKEYS), column=0, padx=5, pady=5, sticky="w")
        tk.Label(root, text="Alternar UI (minimizar/restaurar)").grid(row=len(ASSIGNABLE_HOTKEYS), column=1, padx=5, pady=5, sticky="w")

        btn_manual = tk.Button(root, text="Asignar manualmente", command=self.asignar_manual_global)
        btn_manual.grid(row=len(ASSIGNABLE_HOTKEYS) + 1, column=0, padx=5, pady=10)

        btn_reset = tk.Button(root, text="Resetear todas las teclas", command=self.resetear_hotkeys)
        btn_reset.grid(row=len(ASSIGNABLE_HOTKEYS) + 1, column=1, padx=5, pady=10)

        self.var_inicio = tk.BooleanVar()
        self.var_inicio.set(self.esta_en_inicio())
        chk_inicio = tk.Checkbutton(root, text="Iniciar con Windows", variable=self.var_inicio, command=self.toggle_inicio)
        chk_inicio.grid(row=len(ASSIGNABLE_HOTKEYS) + 2, column=0, padx=5, pady=10, sticky="w")

        self.inicializar_hotkeys()
        keyboard.add_hotkey('f24', self.toggle_window)
        
        # --- CÓDIGO MODIFICADO ---
        # Ocultar la ventana principal al inicio para que se ejecute en segundo plano
        self.root.withdraw()


    def esta_en_inicio(self):
        clave = r"Software\Microsoft\Windows\CurrentVersion\Run"
        try:
            with winreg.OpenKey(winreg.HKEY_CURRENT_USER, clave, 0, winreg.KEY_READ) as reg_key:
                valor, _ = winreg.QueryValueEx(reg_key, "VialMacropad")
                return valor == sys.executable
        except FileNotFoundError:
            return False
        except:
            return False

    def toggle_inicio(self):
        if self.var_inicio.get():
            agregar_al_inicio("VialMacropad", sys.executable)
        else:
            quitar_del_inicio("VialMacropad")

    def inicializar_hotkeys(self):
        for hotkey, ruta in self.config.items():
            if os.path.exists(ruta):
                if hotkey in ASSIGNABLE_HOTKEYS:
                    handler = keyboard.add_hotkey(hotkey, ejecutar_programa, args=[ruta])
                    self.handlers[hotkey] = handler

    def limpiar_hotkeys(self):
        for handler in self.handlers.values():
            keyboard.remove_hotkey(handler)
        self.handlers = {}

    def asignar_manual_global(self):
        def asignar_tecla_manual():
            tecla = entry.get().strip().lower()
            if tecla not in ASSIGNABLE_HOTKEYS:
                messagebox.showerror("Error", f"Tecla inválida. Usa una de: {', '.join(ASSIGNABLE_HOTKEYS)}")
                return
            ruta = filedialog.askopenfilename(title="Seleccionar .exe manualmente", filetypes=[("Ejecutables", "*.exe")])
            if ruta:
                self.asignar_programa_a_tecla(tecla, ruta)
                top.destroy()

        top = tk.Toplevel(self.root)
        top.title("Asignar manualmente")
        tk.Label(top, text=f"Introduce tecla ({', '.join(ASSIGNABLE_HOTKEYS)}):").pack(padx=10, pady=5)
        entry = tk.Entry(top)
        entry.pack(padx=10, pady=5)
        btn_ok = tk.Button(top, text="Asignar", command=asignar_tecla_manual)
        btn_ok.pack(padx=10, pady=10)

    def asignar_programa_desde_lista(self, hotkey, label_widget):
        accesos = obtener_programas_menu_inicio()
        ventana = tk.Toplevel(self.root)
        ventana.title("Elegir programa instalado")
        ventana.geometry("500x500")

        tree = ttk.Treeview(ventana)
        tree["columns"] = ("ruta",)
        tree.heading("#0", text="Programa")
        tree.column("#0", anchor="w")
        tree.heading("ruta", text="Ruta acceso directo")
        tree.column("ruta", anchor="w", width=300)
        tree.pack(expand=True, fill="both", padx=5, pady=5)

        for nombre, ruta_lnk in accesos.items():
            tree.insert("", "end", text=nombre, values=(ruta_lnk,))

        def seleccionar():
            seleccion = tree.selection()
            if not seleccion:
                return
            nombre = tree.item(seleccion[0], "text")
            ruta_lnk = accesos[nombre]
            ruta_real = resolver_acceso_directo(ruta_lnk)
            if ruta_real and os.path.exists(ruta_real):
                self.asignar_programa_a_tecla(hotkey, ruta_real)
                ventana.destroy()
            else:
                messagebox.showerror("Error", f"No se pudo resolver el acceso directo: {ruta_real}")

        btn_seleccionar = tk.Button(ventana, text="Seleccionar", command=seleccionar)
        btn_seleccionar.pack(pady=5)

        btn_manual = tk.Button(ventana, text="Seleccionar manualmente", command=lambda: [ventana.destroy(), self.asignar_manual_global()])
        btn_manual.pack(pady=5)

    def asignar_programa_a_tecla(self, hotkey, ruta):
        self.config[hotkey] = ruta
        guardar_config(self.config)
        self.labels[hotkey].config(text=os.path.basename(ruta))
        if hotkey in self.handlers:
            keyboard.remove_hotkey(self.handlers[hotkey])
            del self.handlers[hotkey]
        handler = keyboard.add_hotkey(hotkey, ejecutar_programa, args=[ruta])
        self.handlers[hotkey] = handler
        print(f"{hotkey} asignado a {ruta}")

    def resetear_hotkeys(self):
        self.limpiar_hotkeys()
        self.config = {}
        guardar_config(self.config)
        for label in self.labels.values():
            label.config(text="(no asignado)")
        messagebox.showinfo("Reset", "Todas las teclas han sido reseteadas.")

    def toggle_window(self):
        if self.root.winfo_viewable():
            self.root.withdraw()
        else:
            self.root.deiconify()
            self.root.lift()
            self.root.focus_force()

if __name__ == "__main__":
    root = tk.Tk()
    app = App(root)
    root.mainloop()
