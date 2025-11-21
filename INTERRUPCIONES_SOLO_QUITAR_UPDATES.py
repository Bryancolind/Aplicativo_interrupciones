import pyodbc
import calendar
import customtkinter as ctk
from tkinter import messagebox, ttk
import pandas as pd
import re
import os
import sys
from datetime import datetime, timedelta
from unidecode import unidecode
import numpy as np
from tkinter import filedialog, Toplevel,Label
from tkcalendar import DateEntry
from dateutil.relativedelta import relativedelta
from datetime import datetime, timedelta
from office365.runtime.auth.authentication_context import AuthenticationContext
from office365.sharepoint.client_context import ClientContext
from office365.sharepoint.client_context import ClientContext
from office365.runtime.auth.authentication_context import AuthenticationContext
from office365.sharepoint.files.file import File
import webbrowser
import locale
from PIL import Image, ImageTk
from PIL import Image
import requests 
ctk.set_appearance_mode("Dark")
ctk.set_default_color_theme("dark-blue")

# --- PARCHE PARA PYINSTALLER Y Tcl/Tk ---
# --- PARCHE MANUAL PARA STORE PYTHON ---
# --- PARCHE PARA RUTAS LOCALES ---
# if getattr(sys, 'frozen', False):
#     os.environ['TCL_LIBRARY'] = os.path.join(sys._MEIPASS, 'mi_tcl')
#     os.environ['TK_LIBRARY'] = os.path.join(sys._MEIPASS, 'mi_tk')
# # ---------------------------------------

# Aquí siguen tus imports (customtkinter, etc...)
class SQLApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Gestor de Eventos")
        #self.root.geometry("1400x800")
        # Obtener dimensiones de la pantalla
        ancho_pantalla = self.root.winfo_screenwidth()
        alto_pantalla = self.root.winfo_screenheight()
        print(ancho_pantalla)
        print(alto_pantalla)
        alto_pantalla = 800
        # Configurar geometría para ventana maximizada
        self.root.geometry(f"{ancho_pantalla}x{alto_pantalla}+0+0")
        self.table_name = "BITACORA_SCADA"
        self.editable_columns = ["Fecha_apertura",'Fecha_cierre',"Tiempo_minutos",'Circuito', 'Subestacion','Ubicacion',"Carga_MVA", "Registro_interrupcion",'Relevador','Observacion'] #columnas que se muestran para editar
        self.editable_invisible_columns =['Fecha_apertura','Fecha_cierre',"Tiempo_horas","Tiempo_minutos", "Carga_MVA",'Registro_interrupcion',"Clasificacion",'Relevador','Interrupcion','Observacion','Usuario_actualizacion','conteo_saifi']#Columnas que se actualizan por debajo del código
        self.display_columns = [
            'Codigo_apertura', 'Codigo_cierre', 'Fecha_apertura', 'Fecha_cierre',
            'Tiempo_horas', 'Tiempo_minutos', 'Zona', 'Sector', 'Subestacion',
            'Circuito', 'Grupo_calidad', 'Tipo_interruptor', 'Equipo_opero','Ubicacion',
            'Carga_MVA', 'Relevador', 'Interrupcion', 'Clasificacion',
            'Registro_interrupcion', 'Observacion','Estado','cambio_hora'
        ] #columnas que se muestran en en aplicativo
        self.float_columns = ['Tiempo_minutos','Grupo_calidad']
        self.conn = None
        self.sort_column = None
        self.sort_order = False  # False = ascendente, True = descendente
        self.hidden_columns = ["conteo_saifi"] #columnas que no se muestran en el aplicativo
        self.filter_windows = {}  # Para mantener ventanas de filtro abiertas
        self.active_filters = {}  # Diccionario de filtros activos {columna: valores}
        self.original_data = pd.DataFrame()  # Copia de los datos sin filtrar

        #Sectores --------------------------------------
        self.editable_columns_2 = ["Fecha_apertura",'Fecha_cierre',"Tiempo_minutos",'Circuito', 'Subestacion','Ubicacion',"Carga_MVA", "Registro_interrupcion",'Relevador','Observacion','Comentario_sector','Revisado_operaciones','Comentario_operaciones']
        self.editable_invisible_columns_2 =['Fecha_apertura','Fecha_cierre',"Tiempo_horas","Tiempo_minutos", "Carga_MVA",'Registro_interrupcion',"Clasificacion",'Relevador','Interrupcion','Observacion','Usuario_actualizacion','Revisado_operaciones','Comentario_operaciones']
        self.display_columns_2 = [
            'Codigo_apertura', 'Codigo_cierre', 'Fecha_apertura', 'Fecha_cierre',
            'Tiempo_horas', 'Tiempo_minutos', 'Sector', 'Subestacion',
            'Circuito', 'Tipo_interruptor', 'Equipo_opero','Ubicacion',
            'Carga_MVA', 'Relevador', 'Interrupcion', 'Clasificacion',
            'Registro_interrupcion', 'Observacion','Revisado_operaciones','Comentario_operaciones','Comentario_sector','Revision','cambio_hora'
        ]      
        self.soporte = 'Si'
        self.share = 'No'


        self.create_login_interface()

    def create_login_interface(self):
        # Crear frame principal
        self.login_frame = ctk.CTkFrame(self.root, fg_color="#2B2B2B")
        self.login_frame.pack(fill="both", expand=True)

# --- Tu función resource_path sigue siendo la misma y es correcta ---
        def resource_path(relative_path):
            """Obtiene ruta válida tanto en .exe como en código normal."""
            try:
                base_path = sys._MEIPASS
            except Exception:
                base_path = os.path.abspath(".")

            return os.path.join(base_path, relative_path)

        # -------------------------------------------------------------------
        
        #1. Cargar la imagen usando Pillow/PIL
        login_img_pil = Image.open(resource_path("assets/login.png"))
        
        #2. Redimensionar (si CustomTkinter lo hacía)
        login_img_resized = login_img_pil.resize((200, 200)) 
        
        # 3. Convertir a un objeto PhotoImage de Tkinter/Pillow
        login_image = ImageTk.PhotoImage(login_img_resized)

        # 4. Usar la imagen en un Label
        # Asegúrate de guardar una referencia a 'login_image' para evitar que sea eliminada por el garbage collector
        
        # Opción A: Usando el Label de CustomTkinter (si el resto de la interfaz lo requiere)
        # Esto debería funcionar si CTkLabel acepta PhotoImage (que sí lo hace)
        self.login_image_ref = login_image # REFERENCIA OBLIGATORIA
        ctk.CTkLabel(self.login_frame, image=login_image, text="").pack(pady=(40, 20))
        
        # Opción B: Usando el Label de Tkinter estándar (si quieres evitar CTk por completo)
        # self.login_image_ref = login_image # REFERENCIA OBLIGATORIA
        # tk.Label(self.login_frame, image=login_image, text="").pack(pady=(40, 20))

        # Título centrado
        ctk.CTkLabel(
            self.login_frame, 
            text="Gestor de Interrupciones - UTCD", 
            font=("Arial", 24, "bold")
        ).pack(pady=(0, 30))

        # Mostrar valores fijos (Servidor y BD)
        server_value = "192.168.100.7"
        database_value = "GestionControl"

        fixed_info = [("Servidor", server_value), ("Base de Datos", database_value)]
        for text, value in fixed_info:
            frame = ctk.CTkFrame(self.login_frame, fg_color="transparent")
            frame.pack(pady=5)
            ctk.CTkLabel(frame, text=text + ":", width=140, anchor="e").pack(side="left", padx=5)
            ctk.CTkLabel(
                frame, 
                text=value, 
                width=300,
                fg_color="#3E3E3E", 
                corner_radius=5, 
                anchor="w"
            ).pack(side="left", padx=5)

        # Campos editables: Usuario y Contraseña
        self.entries = {}
        user_pass_inputs = [("Usuario", "user_entry"), ("Contraseña", "pass_entry")]

        for text, name in user_pass_inputs:
            frame = ctk.CTkFrame(self.login_frame, fg_color="transparent")
            frame.pack(pady=10)

            ctk.CTkLabel(frame, text=text + ":", width=140, anchor="e").pack(side="left", padx=5)

            entry = ctk.CTkEntry(frame, width=300, show="*" if name == "pass_entry" else "")
            entry.pack(side="left")
            entry.bind("<Return>", lambda event: self.validate_and_connect())

            self.entries[name] = entry

            # Toggle mostrar contraseña
            if name == "pass_entry":
                self.show_password = False

                def toggle_password():
                    self.show_password = not self.show_password
                    entry.configure(show="" if self.show_password else "*")
                    toggle_btn.configure(text="🔓" if self.show_password else "🔒")

                toggle_btn = ctk.CTkButton(frame, text="🔒", width=40, command=toggle_password)
                toggle_btn.pack(side="left", padx=5)

        # Botón Conectar centrado
        btn_frame = ctk.CTkFrame(self.login_frame, fg_color="transparent")
        btn_frame.pack(pady=(30, 40))

        ctk.CTkButton(
            btn_frame,
            text="Conectar",
            command=self.validate_and_connect,
            width=200,
            height=45,
            fg_color="#2E8B57",
            hover_color="#245c3d",
            corner_radius=10,
            font=("Arial", 14, "bold")
        ).pack()

        # --- Footer ---
        footer_text = (
            "© UTCD 2025 – Gestión de la Información\n"
            "Sistema para monitoreo, registro y seguimiento de interrupciones eléctricas. \n"
            "Soporte: bryan.colindres@eneeutcd.hn, cristian.umanzor@eneeutcd.hn, ruben.ayestas@eneeutcd.hn \n"
        )

        ctk.CTkLabel(
            self.login_frame,
            text=footer_text,
            font=("Arial", 11),
            justify="center",
            text_color="gray70"
        ).pack(side="bottom", pady=10)

    def validate_and_connect(self):
        self.user = self.entries['user_entry'].get().strip()
        password = self.entries['pass_entry'].get().strip()
        
        if not self.user or not password:
            messagebox.showerror("Error de conexión", "Por favor, complete todos los campos.")
            return
        
        self.connect_to_sql()

    #Función para conectar a sql (base de datos)
    def connect_to_sql(self):
        self.user = self.entries['user_entry'].get().strip()
        password = self.entries['pass_entry'].get().strip()

        conn_str = (
            f"DRIVER={{SQL Server}};"
            f"SERVER=192.168.100.7;"  # Servidor fijo
            f"DATABASE=GestionControl;"  # Base de datos fija
            f"UID={self.user};"
            f"PWD={password}"
        )
        
        try:
            self.current_user = self.user
            self.conn = pyodbc.connect(conn_str)
            self.login_frame.destroy()
            self.create_main_interface()
        except pyodbc.Error as e:
            error_msg = str(e)
            if "Login failed" in error_msg:
                messagebox.showerror("Error de conexión", "Usuario o contraseña incorrectos.")
            else:
                messagebox.showerror("Error de conexión", f"Error desconocido: {error_msg}")

    #FUnción para cortar, copiar pegar texto
    def agregar_menu_contextual(self,entry):
        import tkinter as tk
        """Agrega un menú contextual con opciones de copiar, pegar y cortar."""
        menu = tk.Menu(entry, tearoff=0)  # Usamos Menu de tkinter

        # Copiar al portapapeles
        menu.add_command(label="Copiar", command=lambda: self.copiar(entry))
        # Cortar al portapapeles
        menu.add_command(label="Cortar", command=lambda: self.cortar(entry))
        # Pegar desde el portapapeles
        menu.add_command(label="Pegar", command=lambda: self.pegar(entry))

        def mostrar_menu(event):
            menu.tk_popup(event.x_root, event.y_root)

        entry.bind("<Button-3>", mostrar_menu)

    def copiar(self,entry):
        import tkinter as tk
        # Copiar el texto seleccionado al portapapeles
        entry.clipboard_clear()  # Limpiar el portapapeles
        entry.clipboard_append(entry.get())  # Agregar el texto al portapapeles

    def cortar(self,entry):
        import tkinter as tk
        # Cortar el texto seleccionado y ponerlo en el portapapeles
        self.copiar(entry)  # Primero copiamos
        entry.delete(0, tk.END)  # Luego borramos el contenido del Entry

    def pegar(self,entry):
        import tkinter as tk
        # Pegar el texto del portapapeles
        entry.insert(tk.INSERT, entry.clipboard_get())

    def create_main_interface(self):
        self.main_frame = ctk.CTkFrame(self.root, fg_color="#1E1E1E")
        self.main_frame.pack(fill="both", expand=True, padx=0, pady=0)

        # Contenedor principal con degradado
        main_container = ctk.CTkFrame(self.main_frame, 
                                    fg_color="#2D2D2D",
                                    border_width=0,
                                    corner_radius=0)
        main_container.pack(fill="both", expand=True, padx=20, pady=20)

        # Agregar variables para filtros
        self.filter_codigo = ctk.StringVar()
        self.filter_fecha = ctk.StringVar()

        # Reemplazar la variable existente
        self.filter_fecha_inicio = ctk.StringVar()
        self.filter_fecha_fin = ctk.StringVar()
        # Modificar create_main_interface
        
        # Frame de Filtros ############################################
        filter_frame = ctk.CTkFrame(main_container, fg_color="transparent")
        filter_frame.pack(pady=10, fill="x", padx=20)
        
        # Filtro por Código
        ctk.CTkLabel(filter_frame, text="Filtrar por Código:", width=120).pack(side="left", padx=5)
        codigo_entry = ctk.CTkEntry(
            filter_frame,
            width=200,
            textvariable=self.filter_codigo,
            placeholder_text="Ej: BTE-0001"
        )
        codigo_entry.pack(side="left", padx=5)
        codigo_entry.bind("<KeyRelease>", lambda e: self.apply_filters())
        
        # Filtro por Fecha
        ctk.CTkLabel(filter_frame, text="Filtrar por Fecha:", width=120).pack(side="left", padx=5)
        self.fecha_btn = ctk.CTkButton(
            filter_frame,
            text="Seleccionar Fecha",
            command=self.show_calendar,
            width=150,
            fg_color="#2E8B57",
            hover_color="#245c3d"
        )
        self.fecha_btn.pack(side="left", padx=5)
        
        # En tu frame de filtros, reemplaza el Label existente con:
        ctk.CTkLabel(filter_frame, text="Rango seleccionado:").pack(side="left", padx=5)
        ctk.CTkLabel(filter_frame, textvariable=self.filter_fecha_inicio).pack(side="left", padx=2)
        ctk.CTkLabel(filter_frame, text="a").pack(side="left", padx=2)
        ctk.CTkLabel(filter_frame, textvariable=self.filter_fecha_fin).pack(side="left", padx=2)

        self.filter_mes = ctk.StringVar()  # Variable para el filtro de mes
        # Obtener el mes actual en español
        meses_esp = ["Todos los meses"] + ["Enero", "Febrero", "Marzo", "Abril", "Mayo", "Junio", 
                                      "Julio", "Agosto", "Septiembre", "Octubre", "Noviembre", "Diciembre"]
        self.current_month = meses_esp[datetime.now().month]  # Mes actual en español
        self.current_year = datetime.now().year # Año
        # Filtro por Mes (nuevo)
        ctk.CTkLabel(filter_frame, text="Filtrar por Mes:", width=120).pack(side="left", padx=5)
        mes_menu = ctk.CTkOptionMenu(
            filter_frame,
            variable=self.filter_mes,
            values=meses_esp,
            command=lambda _: self.apply_filters() 
        )
        self.filter_mes.set(self.current_month)  # Establecer el mes actual como valor por defecto
        mes_menu.pack(side="left", padx=5)
        
        # Filtro por Año (nuevo)
        self.filter_ano = ctk.StringVar()  # Variable para el filtro de año
        años = [str(year) for year in range(2025, datetime.now().year + 2)]
        ctk.CTkLabel(filter_frame, text="Filtrar por Año:", width=120).pack(side="left", padx=6)
        ano_menu = ctk.CTkOptionMenu(
            filter_frame,
            variable=self.filter_ano,
            values=años,
            command=lambda _: self.apply_filters() 
        )
        self.filter_ano.set(self.current_year)  # Establecer el mes actual como valor por defecto
        ano_menu.pack(side="left", padx=6)
        
        # Barra de botones estilo moderno
        button_container = ctk.CTkFrame(main_container, 
                                    fg_color="transparent",
                                    height=40)
        button_container.pack(pady=(0, 15), fill="x")

        # Variables para el estado activo
        self.active_tab = ctk.StringVar(value="pendiente")
        self.revision_tab = ctk.StringVar(value="no_revision")
        
        # Función para actualizar el estilo
        def update_button_style(button_active):
            for btn in [pendientes_btn, confirmados_btn]:
                if btn == button_active:
                    btn.configure(fg_color=btn.cget("fg_color"), 
                                font=("Segoe UI Semibold", 13, "bold"),
                                width=190, height = 52)
                else:
                    btn.configure(fg_color=btn.cget("fg_color"), 
                                font=("Segoe UI Semibold", 12),
                                width=160, height = 32)

        # Botones principales
        pendientes_btn = ctk.CTkButton(
            button_container,
            text="🔄 Eventos Pendientes",
            command=lambda: [self.clear_filters(),self.active_tab.set("pendiente"),self.load_table(), update_button_style(pendientes_btn)],
            fg_color="#515A67",
            hover_color="#4B5563",
            corner_radius=6,
            font=("Segoe UI Semibold", 12),
            width=160,
            height=32
        )
        pendientes_btn.pack(side="left", padx=5)

        confirmados_btn = ctk.CTkButton(
            button_container,
            text="✅ Eventos Confirmados",
            command=lambda: [self.clear_filters(),self.active_tab.set("Confirmado"),self.edita_table(), update_button_style(confirmados_btn)],
            fg_color="#10B981",
            hover_color="#086D4B",
            corner_radius=6,
            font=("Segoe UI Semibold", 12),
            width=160,
            height=32
        )
        confirmados_btn.pack(side="left", padx=5)
        

        # Resto de los botones (manteniendo el código original)
        ctk.CTkButton(
            button_container,
            text="✏️ Editar Interrupción",
            command=self.edit_row,
            fg_color="#F59E0B",
            hover_color="#9E6604",
            corner_radius=6,
            font=("Segoe UI Semibold", 12),
            width=160,
            height=32
        ).pack(side="left", padx=5)

        ctk.CTkButton(
            button_container,
            text="❌ Eliminar Interrupción",
            command=self.delete_row,
            fg_color="#EF4444",
            hover_color="#650303",
            corner_radius=6,
            font=("Segoe UI Semibold", 12),
            width=160,
            height=32
        ).pack(side="left", padx=5)
        
        # Añadir botón de actualización
        ctk.CTkButton(
            button_container,
            text="🔄 Actualizar",
            command=self.refresh_table,
            fg_color="#1E1E1E",
            hover_color="#333333",
            border_color="#007ACC",
            border_width=1,
            corner_radius=6,
            font=("Segoe UI Semibold", 12),
            width=100,
            height=32
        ).pack(side="left", padx=5)
        
        ctk.CTkButton(
            button_container,
            text="🔄 Revisión sectores",
            command=lambda: [self.revision_tab.set("si_revision"),self.revision()],
            fg_color="#1E1E1E",
            hover_color="#333333",
            border_color="#007ACC",
            border_width=1,
            corner_radius=6,
            font=("Segoe UI Semibold", 12),
            width=100,
            height=32,
        ).pack(side="left", padx=5)

        # Botón Limpiar Filtros
        ctk.CTkButton(
            button_container,
            text="Limpiar Filtros",
            command=self.clear_filters,
            width=120,
            fg_color="#6c757d",
            hover_color="#5a6268"
        ).pack(side="left", padx=5)
          
        ctk.CTkButton(
            button_container,
            text="📤  Importar datos",
            command=self.importar_excel_y_subir_sql,
            fg_color="#096D27",
            hover_color="#137C08",
            corner_radius=6,
            font=("Segoe UI Semibold", 12),
            width=160,
            height=32
        ).pack(side="right", padx=5)
        
        
        ctk.CTkButton(
            button_container,
            text="📥 Exportar a Excel",
            command=self.export_to_excel,
            fg_color="#3B82F6",
            hover_color="#1F6AE2",
            corner_radius=6,
            font=("Segoe UI Semibold", 12),
            width=160,
            height=32
        ).pack(side="right", padx=5)
        # Contenedor de tabla con sombra moderna
        table_wrapper = ctk.CTkFrame(main_container,
                                fg_color="#333333",
                                corner_radius=8,
                                border_width=0)
        table_wrapper.pack(fill="both", expand=True)

        # Tabla con estilo premium
        self.tree = ttk.Treeview(table_wrapper,
                                style="Fancy.Treeview",
                                show="headings",
                                selectmode="browse")
        
        # Agregar los bindings para copiar filas aquí
        self.tree.bind("<Control-c>", self.copy_row_to_clipboard)
        #self.tree.bind("<Button-3>", self.agregar_menu_contextual)
        
        # Scrollbars personalizadas
        vsb = ctk.CTkScrollbar(
            table_wrapper,
            orientation="vertical",
            command=self.tree.yview,
            button_color="#007ACC",
            button_hover_color="#005F9E",
            fg_color="transparent",
            width=15
        )
        
        hsb = ctk.CTkScrollbar(
            table_wrapper,
            orientation="horizontal",
            command=self.tree.xview,
            button_color="#007ACC",
            button_hover_color="#005F9E",
            fg_color="transparent",
            height=15
        )
        self.root.bind('<F5>', lambda event: self.refresh_table())
        # Configurar estilo de tabla
        self.style = ttk.Style()
        self.style.theme_use("clam")
        
        self.style.configure("Fancy.Treeview",
                            background="#333333",
                            foreground="white",
                            rowheight=36,
                            fieldbackground="#333333",
                            font=("Segoe UI", 11),
                            bordercolor="#444444",
                            borderwidth=0,
                            padding=(8, 4))
        
        self.style.map("Fancy.Treeview",
                    background=[('selected', '#007ACC')],
                    foreground=[('selected', 'white')])
        
        self.style.configure("Fancy.Treeview.Heading",
                            background="#444444",
                            foreground="white",
                            font=("Segoe UI Semibold", 11),
                            relief="flat",
                            padding=(12, 8),
                            borderwidth=0,
                            anchor="center")
        
        self.style.map("Fancy.Treeview.Heading",
                    background=[('active', '#555555')])

        # Layout profesional
        self.tree.grid(row=0, column=0, sticky="nsew", padx=2, pady=2)
        vsb.grid(row=0, column=1, sticky="ns", padx=2, pady=2)
        hsb.grid(row=1, column=0, sticky="ew", padx=2, pady=2)
        
        table_wrapper.grid_rowconfigure(0, weight=1)
        table_wrapper.grid_columnconfigure(0, weight=1)
        self.tree.configure(yscrollcommand=vsb.set, xscrollcommand=hsb.set)

        # Configurar columnas
        self.tree["columns"] = self.display_columns
        for col in self.display_columns:
            self.tree.heading(col, 
                            text=col.upper(), 
                            command=lambda c=col: self.sort_treeview(c),
                            anchor="center")
            self.tree.column(col, 
                            width=180, 
                            anchor="center", 
                            minwidth=140,
                            stretch=False)

        # Efectos de scroll fluidos
        self.tree.bind("<MouseWheel>", self.on_vertical_scroll)
        self.tree.bind("<Shift-MouseWheel>", self.on_horizontal_scroll)

        # Contador de registros estilo dashboard
        self.record_count = ctk.CTkLabel(main_container,
                                    text="⚡ Registros Activos: 0",
                                    font=("Segoe UI Semibold", 12),
                                    text_color="#007ACC",
                                    anchor="w",
                                    height=20)
        self.record_count.pack(pady=(10, 0), padx=5, fill="x")
        
        # Agregar el menú contextual
        self.agregar_menu_contextual(codigo_entry)
        # Efecto hover premium
        self.tree.tag_configure('hover', background='#3A3A3A')
        self.tree.bind("<Motion>", self.on_hover)

        self.load_table()
        
        # Estilo inicial
        update_button_style(pendientes_btn)
        self.load_table()
    
    def revision(self):
        self.main_frame.destroy()
        self.main_frame = ctk.CTkFrame(self.root, fg_color="#1E1E1E")
        self.main_frame.pack(fill="both", expand=True, padx=0, pady=0)

        # Contenedor principal con degradado
        main_container = ctk.CTkFrame(self.main_frame, 
                                    fg_color="#2D2D2D",
                                    border_width=0,
                                    corner_radius=0)
        main_container.pack(fill="both", expand=True, padx=20, pady=20)

        # Agregar variables para filtros
        self.filter_codigo = ctk.StringVar()
        self.filter_fecha = ctk.StringVar()

        # Reemplazar la variable existente
        self.filter_fecha_inicio = ctk.StringVar()
        self.filter_fecha_fin = ctk.StringVar()
        # Modificar create_main_interface
        
        # Frame de Filtros ############################################
        filter_frame = ctk.CTkFrame(main_container, fg_color="transparent")
        filter_frame.pack(pady=10, fill="x", padx=20)
        
        # Filtro por Código

        ctk.CTkButton(
            filter_frame,
            text="🔙 Volver",
            command=lambda: [self.revision_tab.set("no_revision"),self.main_interface_nueva()],
            fg_color="#045b8c",
            hover_color="#053c5b",
            corner_radius=6,
            font=("Segoe UI Semibold", 12),
            width=160,
            height=32
        ).pack(side="left", padx=5)
        
        # Filtro por Fecha
        ctk.CTkLabel(filter_frame, text="Filtrar por Fecha:", width=120).pack(side="left", padx=5)
        self.fecha_btn = ctk.CTkButton(
            filter_frame,
            text="Seleccionar Fecha",
            command=self.show_calendar,
            width=150,
            fg_color="#2E8B57",
            hover_color="#245c3d"
        )
        self.fecha_btn.pack(side="left", padx=5)
        
        # En tu frame de filtros, reemplaza el Label existente con:
        ctk.CTkLabel(filter_frame, text="Rango seleccionado:").pack(side="left", padx=5)
        ctk.CTkLabel(filter_frame, textvariable=self.filter_fecha_inicio).pack(side="left", padx=2)
        ctk.CTkLabel(filter_frame, text="a").pack(side="left", padx=2)
        ctk.CTkLabel(filter_frame, textvariable=self.filter_fecha_fin).pack(side="left", padx=2)

        self.filter_mes = ctk.StringVar()  # Variable para el filtro de mes
        # Obtener el mes actual en español
        meses_esp = ["Todos los meses"] + ["Enero", "Febrero", "Marzo", "Abril", "Mayo", "Junio", 
                                      "Julio", "Agosto", "Septiembre", "Octubre", "Noviembre", "Diciembre"]
        self.current_month = meses_esp[datetime.now().month]  # Mes actual en español
        # Filtro por Mes (nuevo)
        ctk.CTkLabel(filter_frame, text="Filtrar por Mes:", width=120).pack(side="left", padx=5)
        mes_menu = ctk.CTkOptionMenu(
            filter_frame,
            variable=self.filter_mes,
            values=meses_esp,
            command=lambda _: self.aplicar_filtro() 
        )
        self.filter_mes.set(self.current_month)  # Establecer el mes actual como valor por defecto
        mes_menu.pack(side="left", padx=5)
        
        # Filtro por Año (nuevo)
        self.filter_ano = ctk.StringVar()  # Variable para el filtro de año
        años = [str(year) for year in range(2025, datetime.now().year + 2)]
        ctk.CTkLabel(filter_frame, text="Filtrar por Año:", width=120).pack(side="left", padx=6)
        ano_menu = ctk.CTkOptionMenu(
            filter_frame,
            variable=self.filter_ano,
            values=años,
            command=lambda _: self.apply_filters() 
        )
        self.filter_ano.set(self.current_year)  # Establecer el mes actual como valor por defecto
        ano_menu.pack(side="left", padx=6)
        
        # Botón Limpiar Filtros
        ctk.CTkButton(
            filter_frame,
            text="Limpiar Filtros",
            command=self.limpiar_filtros,
            width=120,
            fg_color="#6c757d",
            hover_color="#5a6268"
        ).pack(side="right", padx=5)

        # Barra de botones estilo moderno
        self.button_container = ctk.CTkFrame(main_container, 
                                    fg_color="transparent",
                                    height=40)
        self.button_container.pack(pady=(0, 15), fill="x")

        # Variables para el estado activo
        self.active_tab_2 = ctk.StringVar(value="confirmado")
        
        # Función para actualizar el estilo
        def update_button_style(button_active):
            for btn in [confirmados_btn]:
                if btn == button_active:
                    btn.configure(fg_color=btn.cget("fg_color"), 
                                font=("Segoe UI Semibold", 13, "bold"),
                                width=190, height = 52)
                else:
                    btn.configure(fg_color=btn.cget("fg_color"), 
                                font=("Segoe UI Semibold", 12),
                                width=160, height = 32)

        # Botones principales

        confirmados_btn = ctk.CTkButton(
            self.button_container,
            text="✅ Registros revisados",
            command=lambda: [self.active_tab_2.set("Confirmado"),self.revisados(), update_button_style(confirmados_btn)],
            fg_color="#45026d",
            hover_color="#370257",
            corner_radius=6,
            font=("Segoe UI Semibold", 12),
            width=160,
            height=32
        )
        confirmados_btn.pack(side="left", padx=5)
        
        
        ctk.CTkButton(
            self.button_container,
            text="Ampliar",
            command=self.ampliar,
            fg_color="#4b2f4d",
            hover_color="#29760a",
            corner_radius=6,
            font=("Segoe UI Semibold", 12),
            width=160,
            height=32
        ).pack(side="left", padx=5)
        
        # Añadir botón de actualización
        ctk.CTkButton(
            self.button_container,
            text="🔄 Actualizar",
            command=self.actualizar,
            fg_color="#1E1E1E",
            hover_color="#333333",
            border_color="#007ACC",
            border_width=1,
            corner_radius=6,
            font=("Segoe UI Semibold", 12),
            width=100,
            height=32
        ).pack(side="left", padx=5)
        
        ctk.CTkButton(
            self.button_container,
            text="📥 Exportar a Excel",
            command=self.export_to_excel_revision,
            fg_color="#4A90E2",
            hover_color="#357ABD",
            corner_radius=6,
            font=("Segoe UI Semibold", 12),
            width=160,
            height=32
        ).pack(side="right", padx=5)
        # Contenedor de tabla con sombra moderna
        table_wrapper = ctk.CTkFrame(main_container,
                                fg_color="#333333",
                                corner_radius=8,
                                border_width=0)
        table_wrapper.pack(fill="both", expand=True)

        # Tabla con estilo premium
        self.tree = ttk.Treeview(table_wrapper,
                                style="Fancy.Treeview",
                                show="headings",
                                selectmode="browse")
        
        # Agregar los bindings para copiar filas aquí
        self.tree.bind("<Control-c>", self.copy_row_to_clipboard)
        #self.tree.bind("<Button-3>", self.agregar_menu_contextual)


        # Scrollbars personalizadas
        vsb = ctk.CTkScrollbar(
            table_wrapper,
            orientation="vertical",
            command=self.tree.yview,
            button_color="#007ACC",
            button_hover_color="#005F9E",
            fg_color="transparent",
            width=15
        )
        
        hsb = ctk.CTkScrollbar(
            table_wrapper,
            orientation="horizontal",
            command=self.tree.xview,
            button_color="#007ACC",
            button_hover_color="#005F9E",
            fg_color="transparent",
            height=15
        )
        self.root.bind('<F5>', lambda event: self.actualizar())
        # Configurar estilo de tabla
        self.style = ttk.Style()
        self.style.theme_use("clam")
        
        # Estilo global para el Treeview
        self.style.configure("Fancy.Treeview",
                            background="#333333",  # Fondo oscuro para la tabla
                            foreground="white",
                            rowheight=36,
                            fieldbackground="#333333",  # Fondo oscuro también en las celdas
                            font=("Segoe UI", 11),
                            bordercolor="#444444",
                            borderwidth=0,
                            padding=(8, 4))

        # Asegúrate de que el fondo seleccionado no afecte los tags
        self.style.map("Fancy.Treeview",
            background=[('selected', '#007ACC')],
            foreground=[('selected', 'white')]
        )

        # Configuración específica de colores para las filas etiquetadas
        self.tree.tag_configure("verde", background="#1c7e02", foreground="white")  # verde claro
        self.tree.tag_configure("rojo", background="lightcoral")     # rojo claro
        self.tree.tag_configure("naranja", background="orange")      # anaranjado claro


        
        self.style.configure("Fancy.Treeview.Heading",
                            background="#444444",
                            foreground="white",
                            font=("Segoe UI Semibold", 11),
                            relief="flat",
                            padding=(12, 8),
                            borderwidth=0,
                            anchor="center")
        
        self.style.map("Fancy.Treeview.Heading",
                    background=[('active', '#555555')])

        # Layout profesional
        self.tree.grid(row=0, column=0, sticky="nsew", padx=2, pady=2)
        vsb.grid(row=0, column=1, sticky="ns", padx=2, pady=2)
        hsb.grid(row=1, column=0, sticky="ew", padx=2, pady=2)
        
        table_wrapper.grid_rowconfigure(0, weight=1)
        table_wrapper.grid_columnconfigure(0, weight=1)
        self.tree.configure(yscrollcommand=vsb.set, xscrollcommand=hsb.set)

        # Configurar columnas
        self.tree["columns"] = self.display_columns_2
        for col in self.display_columns_2:
            self.tree.heading(col, 
                            text=col.upper(), 
                            command=lambda c=col: self.sort_treeview(c),
                            anchor="center")
            self.tree.column(col, 
                            width=180, 
                            anchor="center", 
                            minwidth=140,
                            stretch=False)

        # Efectos de scroll fluidos
        self.tree.bind("<MouseWheel>", self.on_vertical_scroll)
        self.tree.bind("<Shift-MouseWheel>", self.on_horizontal_scroll)

        # Contador de registros estilo dashboard
        self.record_count = ctk.CTkLabel(main_container,
                                    text="⚡ Registros Activos: 0",
                                    font=("Segoe UI Semibold", 12),
                                    text_color="#007ACC",
                                    anchor="w",
                                    height=20)
        self.record_count.pack(pady=(10, 0), padx=5, fill="x")
        
        # Agregar el menú contextual
        # Efecto hover premium
        self.tree.tag_configure('hover', background='#3A3A3A')
        self.tree.bind("<Motion>", self.color_mouse)

        # Estilo inicial
        update_button_style(confirmados_btn)
        self.revisados()

    def main_interface_nueva(self):
        self.main_frame.destroy()
        self.create_main_interface()
    
    
    def revisados(self): 
        
        # Verifica si las etiquetas con círculos ya existen en el contenedor
        existing_labels = [child for child in self.button_container.winfo_children() if isinstance(child, ctk.CTkCanvas)]
        
        # Si no existen, las agrega
        if not existing_labels:
            self.create_label_with_circle(self.button_container, "Pendiente", "#0a025d")  # Naranja
            self.create_label_with_circle(self.button_container, "Aceptado", "#1c7e02")  # Verde
            self.create_label_with_circle(self.button_container, "Rechazado", "#ab0808")  # Rojo
    
        try:
            two_months_ago = datetime.now().replace(day=1).date().strftime('%Y-%m-%d')
            base_query = f"""
                SELECT {', '.join(f'[{col}]' for col in self.display_columns_2)} 
                FROM {self.table_name}
                WHERE Estado = 'Confirmado' and Revisado_Operaciones in ('Pendiente','Aceptado','Pendiente-sin soporte','Rechazado')
                AND Activo = 1
            """

            params = []
            conditions = []

            
            # Filtro de fecha: si hay un valor en filter_fecha, lo agregamos a las condiciones
            if self.filter_fecha.get():
                conditions.append("CAST(Fecha_apertura AS DATE) = ?")
                params.append(self.filter_fecha.get())

            # Filtro de código: si hay un valor en filter_codigo, lo agregamos a las condiciones
            if self.filter_codigo.get():
                conditions.append("Codigo_apertura LIKE ?")
                params.append(f"%{self.filter_codigo.get()}%")
            

            
            # Lista de meses en español
            meses_esp = [
                "Todos los meses", "enero", "febrero", "marzo", "abril", "mayo", "junio", 
                "julio", "agosto", "septiembre", "octubre", "noviembre", "diciembre"
            ]
            
            # Aquí usas la lista en español para obtener el número del mes
            selected_month = self.filter_mes.get().lower()  # Convertir a minúsculas
            if selected_month != "todos los meses":
                # Convertir el nombre del mes en español a su número correspondiente
                month_number = meses_esp.index(selected_month)  # Esto devuelve el índice del mes
                conditions.append("MONTH(Fecha_apertura) = ?")
                params.append(month_number)
                
            selected_year = self.filter_ano.get()
            if selected_year:
                conditions.append("YEAR(Fecha_apertura) = ?")
                params.append(int(selected_year))
            # Si hay condiciones de filtro, agregarlas a la consulta base
            if conditions:
                base_query += " AND " + " AND ".join(conditions)

            # Ordenar por fecha de apertura y cierre
            base_query += " ORDER BY CAST(SUBSTRING(Revisado_sector, CHARINDEX('-', Revisado_sector, CHARINDEX('-', Revisado_sector) + 1) + 2, 19) AS DATEtime) asc"
            
            print(base_query)
            # Ejecutar la consulta y cargar los resultados en un DataFrame
            df = pd.read_sql(base_query, self.conn, params=params)
            print("Columnas cargadas:", df.columns.tolist())
            # Convertir las columnas de fecha
            date_columns = [col for col in df.columns if 'Fecha' in col]
            for col in date_columns:
                df[col] = pd.to_datetime(df[col], errors='coerce')
            
            # Procesar columnas numéricas
            df = df.replace({np.nan: None, '': None,' ':None})
            for col in self.float_columns:
                if col in df.columns:
                    df[col] = df[col].fillna(0).apply(lambda x: round(float(x))).astype(int)
            df['Tiempo_horas'] = df['Tiempo_horas'].apply(lambda x: np.format_float_positional(float(x), precision=2, unique=False, fractional=False, trim='k') if pd.notna(x) else x)
            
            
            
            # Mostrar la tabla
            self.mostrar_tabla(df)
            self.record_count.configure(text=f"Registros: {len(df)}")
            return df
        except Exception as e:
            print(f'error esta enn{e}')
            self.handle_error(e)

    #boton de ampliar para ver las cosas
    def ampliar(self):
        selected = self.tree.selection()
        if not selected:
            messagebox.showwarning("Advertencia", "Seleccione un registro para editar")
            return
        self.soporte == 'Si'
        self.share = 'No'
        self.item = selected[0]
        values = self.tree.item(self.item)["values"]
        data = dict(zip(self.display_columns_2, values))

        self.codigo_apertura = data['Codigo_apertura']
        self.codigo_cierre =   data['Codigo_cierre']
        

        # Establece el locale en español para que los nombres de los meses salgan en español
        locale.setlocale(locale.LC_TIME, 'Spanish_Spain.1252')  # En Windows

        # Suponiendo que data['Fecha_apertura'] es un string
        fecha_str = data['Fecha_apertura'][:19]  # Recorta los microsegundos extra
        fecha_dt = datetime.strptime(fecha_str, "%Y-%m-%d %H:%M:%S")

        # Extraer el mes en palabras
        mes_en_palabras = fecha_dt.strftime('%B')  # Por ejemplo, 'marzo'

        self.fecha_mes = mes_en_palabras        
        
        if 'Exclusión' in data['Revision'].lower() or 'exclusión' in data['Revision'].lower():
            print('aqui')
            self.share = 'Si'
        self.sector_ = data["Sector"]
        self.edit_window = ctk.CTkToplevel(self.root)
        self.edit_window.title("Editar Registro")
        self.edit_window.geometry("620x720")
        self.edit_window.grab_set()

        entries = {}

        sectores = ['TEGUCIGALPA-DANLI','CHOLUTECA','COMAYAGUA',
                    'JUTICALPA','TOCOA-LA CEIBA','SAN PEDRO SULA',
                    'VILLANUEVA','SANTA ROSA DE COPAN',
                    'EL PROGRESO-SANTA CRUZ DE YOJOA']
        
        encontrado = [s for i, s in enumerate(sectores) if self.sector_  in s]
        
        self.carpetas= data['Sector']


        selected = self.tree.selection()
        if not selected:
            messagebox.showwarning("Advertencia", "Seleccione un registro para editar")
            return
        

        registros_info = self.obtener_registros_interrupcion()
        
        display_values = [f"{r[0]} - {r[1]}" for r in registros_info]
        registro_mapping = {display: r[0] for display, r in zip(display_values, registros_info)}

        relevador_mapping = {"SIN INTERRUPCIÓN": "SIN INTERRUPCIÓN", "APERTURA REMOTA": "APERTURA REMOTA", "APERTURA LOCAL": "APERTURA LOCAL", "FALLA EN SIN": "FALLA EN SIN", "51ABCN": "51ABCN", "51ABC": "51ABC", "51ABN": "51ABN", "51ACN": "51ACN", "51BCN": "51BCN", "51AN": "51AN", "51BN": "51BN", "51CN": "51CN", "51AB": "51AB", "51AC": "51AC", "51BC": "51BC", "51A": "51A", "51B": "51B", "51C": "51C", "51N": "51N", "51FN": "51FN", "51F": "51F", "21 (DISTANCIA)": "21 (DISTANCIA)", "27 (BAJO VOLTAJE)": "27 (BAJO VOLTAJE)", "50 (SOBRE CORRIENTE INSTANTÁNEO)": "50 (SOBRE CORRIENTE INSTANTÁNEO)", "51 (SOBRE CORRIENTE TEMPORIZADO)": "51 (SOBRE CORRIENTE TEMPORIZADO)", "63 (PRESIÓN)": "63 (PRESIÓN)", "67 (DIRECCIONAL DE SOBRE CORRIENTE)": "67 (DIRECCIONAL DE SOBRE CORRIENTE)", "79 (RECIERRE)": "79 (RECIERRE)", "81 (FRECUENCIA)": "81 (FRECUENCIA)", "86 (BLOQUEO)": "86 (BLOQUEO)", "87 (DIFERENCIAL)": "87 (DIFERENCIAL)", "90 (REGULACIÓN)": "90 (REGULACIÓN)", "NO INDICO": "NO INDICO"}

        relevador_values = [
            "SIN INTERRUPCIÓN", "APERTURA REMOTA", "APERTURA LOCAL", "FALLA EN SIN",
            "51ABCN", "51ABC", "51ABN", "51ACN", "51BCN", "51AN", "51BN", "51CN", 
            "51AB", "51AC", "51BC", "51A", "51B", "51C", "51N", "51FN", "51F", 
            "21 (DISTANCIA)", "27 (BAJO VOLTAJE)", "50 (SOBRE CORRIENTE INSTANTÁNEO)", 
            "51 (SOBRE CORRIENTE TEMPORIZADO)", "63 (PRESIÓN)", "67 (DIRECCIONAL DE SOBRE CORRIENTE)", 
            "79 (RECIERRE)", "81 (FRECUENCIA)", "86 (BLOQUEO)", "87 (DIFERENCIAL)", 
            "90 (REGULACIÓN)", "NO INDICO"
        ]

        
        relevador_estado = {"Aceptado":"Aceptado",'Rechazado':'Rechazado','Pendiente':'Pendiente','Pendiente-sin soporte':'Pendiente-sin soporte'}
        estados = ['Aceptado','Rechazado','Pendiente','Pendiente-sin soporte']
        class AutoCompleteCombobox(ctk.CTkComboBox):
            def __init__(self, master, values, **kwargs):
                super().__init__(master, values=values, **kwargs)
                self.full_values = values
                self.filtered_values = values
                self.dropdown = None
                self._entry.bind("<KeyRelease>", self.update_filter)
                self._entry.bind("<FocusOut>", lambda e: self.close_dropdown(delay=100))
                self._entry.bind("<Down>", self.open_dropdown)

            def update_filter(self, event=None):
                input_text = self._entry.get().lower()
                self.filtered_values = [v for v in self.full_values if input_text in v.lower()]
                
                if self.dropdown and self.dropdown.winfo_exists():
                    self.update_dropdown()
                else:
                    self.open_dropdown()
                
                self.configure(values=self.filtered_values)

            def open_dropdown(self, event=None):
                if not self.filtered_values:
                    return
                    
                if self.dropdown is None or not self.dropdown.winfo_exists():
                    x = self.winfo_rootx()
                    y = self.winfo_rooty() + self.winfo_height()
                    width = self.winfo_width()
                    
                    self.dropdown = ctk.CTkToplevel(self)
                    self.dropdown.overrideredirect(True)
                    self.dropdown.geometry(f"{width}x200+{x}+{y}")
                    self.dropdown.attributes("-topmost", True)
                    
                    self.scroll_frame = ctk.CTkScrollableFrame(self.dropdown)
                    self.scroll_frame.pack(fill="both", expand=True)
                    self.update_dropdown()
                    
                    self.dropdown.bind("<Button-1>", self.check_click_outside)

            def update_dropdown(self):
                for widget in self.scroll_frame.winfo_children():
                    widget.destroy()
                
                for value in self.filtered_values:
                    btn = ctk.CTkButton(
                        self.scroll_frame,
                        text=value,
                        width=self.winfo_width(),
                        anchor="w",
                        command=lambda v=value: self.select_value(v),
                        fg_color="transparent",
                        hover_color="#3B8ED0",
                        text_color="#FFFFFF",
                        font=("Arial", 12)
                    )
                    btn.pack(pady=1, padx=0)

            def select_value(self, value):
                self.set(value)
                if self.dropdown:
                    self.dropdown.destroy()
                    self.dropdown = None
                self._entry.focus()

            def check_click_outside(self, event):
                if self.dropdown and self.dropdown.winfo_exists():
                    self.dropdown.update_idletasks()
                    d_x = self.dropdown.winfo_x()
                    d_y = self.dropdown.winfo_y()
                    d_width = self.dropdown.winfo_width()
                    d_height = self.dropdown.winfo_height()
                    
                    if not (d_x <= event.x_root <= d_x + d_width and
                            d_y <= event.y_root <= d_y + d_height):
                        self.close_dropdown(delay=0)

            def close_dropdown(self, delay=0):
                if self.dropdown and self.dropdown.winfo_exists():
                    if delay > 0:
                        self.dropdown.after(delay, self.dropdown.destroy)
                    else:
                        self.dropdown.destroy()
                    self.dropdown = None
        
        # Creación de campos de edición
        read_only_fields = ['Circuito', 'Subestacion', 'Ubicacion']

        # Crear scrollable frame
        scrollable_frame = ctk.CTkScrollableFrame(self.edit_window, width=600, height=650)
        scrollable_frame.pack(padx=10, pady=10, fill="both", expand=True)
        
        for col in self.editable_columns_2:
            frame = ctk.CTkFrame(scrollable_frame, fg_color="transparent")
            frame.pack(pady=8, fill="x", padx=15)
            
            ctk.CTkLabel(frame, text=f"{col}:", width=140, anchor="e").pack(side="left", padx=5)

            if col in read_only_fields:
                value = data.get(col, "")
                label = ctk.CTkLabel(
                    frame,
                    text=value,
                    width=300,
                    fg_color="#343638",
                    corner_radius=3,
                    anchor="w"
                )
                label.pack(side="left")
                continue  # Saltar a la siguiente iteración
            
            # Campo Relevador con lista desplegable
            # Campo Relevador con lista desplegable
            elif col == "Relevador":
                entry = AutoCompleteCombobox(
                    frame, 
                    values=relevador_values, 
                    width=300,
                    dropdown_fg_color="#2B2B2B",
                    button_color="#4A4A4A",
                    border_color="#565B5E",
                    fg_color="#343638",
                    text_color="#FFFFFF"
                )
                
                current_relevador = str(data.get(col, ""))
                #print(f"Valor en data['Relevador']: {current_relevador}")  # Depuración
                current_display = next(
                    (disp for disp, reg in relevador_mapping.items() if reg == current_relevador),
                    ""
                )
                
                entry.set(current_display)
                entry.relevador_mapping = relevador_mapping
                
            # Campo Relevador con lista desplegable
            elif col == "Revisado_operaciones":
                entry = AutoCompleteCombobox(
                    frame, 
                    values=estados, 
                    width=300,
                    dropdown_fg_color="#2B2B2B",
                    button_color="#4A4A4A",
                    border_color="#565B5E",
                    fg_color="#343638",
                    text_color="#FFFFFF"
                )
                
                current_relevador = str(data.get(col, ""))
                #print(f"Valor en data['operaciones']: {current_relevador}")  # Depuración
                current_display = next(
                    (disp for disp, reg in relevador_estado.items() if reg == current_relevador),
                    ""
                )
                
                entry.set(current_display)
                entry.relevador_estado = relevador_estado

                
            # Campo Minutos con botón de cálculo
            elif col == "Tiempo_minutos":
                entry_frame = ctk.CTkFrame(frame, fg_color="transparent")
                entry_frame.pack(side="left")
                
                entry = ctk.CTkEntry(
                    entry_frame,
                    width=250,
                    border_width=1,
                    fg_color="#343638",
                    text_color="#FFFFFF"
                )
                entry.insert(0, str(data.get(col, "")))
                entry.pack(side="left", padx=(0, 5))

                # Capturar la referencia correcta del entry usando lambda
                def crear_calcular_minutos(entry_widget):
                    def calcular_minutos():
                        try:
                            fecha_apertura = datetime.strptime(
                                entries["Fecha_apertura"].get(), 
                                "%Y-%m-%d %H:%M:%S"
                            )
                            fecha_cierre = datetime.strptime(
                                entries["Fecha_cierre"].get(), 
                                "%Y-%m-%d %H:%M:%S"
                            )
                            diferencia = fecha_cierre - fecha_apertura
                            minutos = int(diferencia.total_seconds() / 60+0.5)
                            entry_widget.delete(0, "end")
                            entry_widget.insert(0, str(minutos))
                        except Exception as e:
                            messagebox.showerror(
                                "Error", 
                                f"Error calculando minutos: Verifique los formatos de fecha\nDetalle: {str(e)}"
                            )
                    return calcular_minutos

                ctk.CTkButton(
                    entry_frame,
                    text="🕒 Calcular",
                    width=80,
                    height=28,
                    command=crear_calcular_minutos(entry),  # Pasamos la referencia correcta
                    fg_color="#4A752C",
                    hover_color="#5A8C3F",
                    font=("Arial", 10)
                ).pack(side="left")

            elif col == "Registro_interrupcion":
                entry = AutoCompleteCombobox(
                    frame, 
                    values=display_values, 
                    width=300,
                    dropdown_fg_color="#2B2B2B",
                    button_color="#4A4A4A",
                    border_color="#565B5E",
                    fg_color="#343638",
                    text_color="#FFFFFF"
                )
                current_registro = str(data.get(col, ""))
                current_display = next(
                    (disp for disp, reg in registro_mapping.items() if reg == current_registro),
                    ""
                )
                entry.set(current_display)
                entry.registro_mapping = registro_mapping
                
            elif col in ["Observacion", "Comentario_operaciones",'Comentario_sector']:
                entry = ctk.CTkTextbox(
                    frame, 
                    width=300, 
                    height=180, 
                    border_width=1,
                    fg_color="#343638",
                    text_color="#FFFFFF",
                    wrap="word"
                )
                entry.insert("1.0", data.get(col, ""))
                self.agregar_menu_contextual(entry)
                
                def handle_return(event):
                    event.widget.insert("insert", "\n")
                    return "break"
                    
                entry.bind("<Return>", handle_return)
                entry.bind("<KeyRelease-Return>", lambda e: "break")
                
            else:
                entry = ctk.CTkEntry(
                    frame, 
                    width=300, 
                    border_width=1,
                    fg_color="#343638",
                    text_color="#FFFFFF"
                )
                entry.insert(0, str(data.get(col, "")))


            entry.pack(side="left")
            entries[col] = entry

        # Botones de acción dentro del scroll
        btn_frame = ctk.CTkFrame(scrollable_frame, fg_color="transparent")
        btn_frame.pack(pady=20)

        ctk.CTkButton(
            btn_frame,
            text="Guardar cambios",
            command=lambda: self.guardar(self.item, entries, self.edit_window),
            fg_color="#41b612",
            hover_color="#29760a",
            corner_radius=6,
            font=("Segoe UI Semibold", 12),
            width=160,
            height=32
        ).pack(side="left", padx=5)

        ctk.CTkButton(
            btn_frame,
            text="Cancelar",
            command=self.edit_window.destroy,
            width=120,
            height=35,
            fg_color="#dc3545",
            hover_color="#c82333",
            font=("Arial", 12)
        ).pack(side="left", padx=15)
        
        if self.share == 'Si':
            ctk.CTkButton(
                btn_frame,
                text="Sharepoint",
                command=lambda: self.abrir_si_existe_sharepoint(self.codigo_apertura, self.codigo_cierre, self.carpetas,self.fecha_mes),
                width=120,
                height=35,
                fg_color="#057189",
                hover_color="#c82333",
                font=("Arial", 12)
            ).pack(side="left", padx=15)

    def guardar(self, item, entries, window):
        original_values = self.tree.item(item)["values"]
        original_data = dict(zip(self.display_columns_2, original_values))
        
        # Campos de solo lectura
        read_only_fields = ['Circuito', 'Subestacion', 'Ubicacion']
        # Construir new_data excluyendo campos bloqueados
        new_data = {}
        for col in self.editable_columns_2:
            if col in read_only_fields:
                continue  # Saltar campos no editables
                
            if col == "Registro_interrupcion":
                # Obtener valor real del registro
                selected_display = entries[col].get().strip()
                registro_mapping = getattr(entries[col], 'registro_mapping', {})
                registro_numero = registro_mapping.get(selected_display, "")
                
                if not registro_numero and selected_display:
                    messagebox.showerror("Error", "Selección inválida para Registro de Interrupción")
                    return
                    
                new_data[col] = registro_numero
                
            elif col in ["Observacion", "Comentario_operaciones",'Comentario_sector']:
                new_data[col] = entries[col].get("1.0", "end-1c").strip()
                
            else:
                new_data[col] = entries[col].get().strip()
        
        fecha_revision = str(datetime.now().strftime('%Y-%m-%d %H:%M:%S'))
        new_data['Usuario_actualizacion'] = f'{self.user} - {fecha_revision}'

        # 1. Validación de tipos de datos
        type_mapping = {
            'int': int,
            'varchar': str,
            'decimal': float,
            'date': lambda x: datetime.strptime(x, '%Y-%m-%d').date(),
            'bit': lambda x: bool(int(x)),
            'datetime': datetime.fromisoformat
        }

        try:
            # Calcular tiempo automático
            if 'Tiempo_minutos' in new_data:
                total_minutos = int(new_data['Tiempo_minutos'])
                horas = total_minutos / 60
                minutos = total_minutos
                new_data['Tiempo_horas'] = horas
                new_data['Tiempo_minutos'] = minutos

            # Obtener clasificación
            if 'Registro_interrupcion' in new_data and new_data['Registro_interrupcion']:
                cursor = self.conn.cursor()
                query = """
                    SELECT [Clasificacion_a], Descripcion_falla
                    FROM [GestionControl].[dbo].[CLASIFICACION_INTERRUPCIONES]
                    WHERE Registro = ?
                """
                cursor.execute(query, (new_data['Registro_interrupcion'],))
                result = cursor.fetchall()
                
                if result:
                    new_data['Clasificacion'] = result[0][0]
                    new_data['Interrupcion'] = result[0][1]

                else:
                    messagebox.showerror('Error', "No se encontró la clasificación para el registro seleccionado")
                    return
                cursor.close()

        except ValueError as e:
            messagebox.showerror("Error", f"Error de conversión de datos: {e}")
            return
        except Exception as e:
            messagebox.showerror("Error", f"Error inesperado: {e}")
            return

        if new_data['Tiempo_horas'] < 0.06:
            new_data['Registro_interrupcion'] = 17
            new_data['Interrupcion'] = 'INSTANTANEA'
            new_data['Clasificacion'] = 'E'
        # Validación de tipos de datos
        type_errors = []
        for col in self.editable_columns_2:
            if col in read_only_fields:
                continue  # Saltar validación para campos bloqueados
                
            sql_type = self.get_sql_type(col).lower()
            base_type = sql_type.split('(')[0].strip()
            converter = type_mapping.get(base_type, str)
            
            try:
                if new_data.get(col, '') == '':
                    new_data[col] = None if base_type != 'varchar' else ''
                else:
                    new_data[col] = converter(new_data[col])
            except Exception as e:
                type_errors.append(f"Columna {col}: Valor '{new_data[col]}' no válido para tipo {sql_type}")

        if type_errors:
            messagebox.showerror("Error de tipos", "\n".join(type_errors))
            return

        # Verificar cambios excluyendo campos bloqueados
        has_changes = any(
            str(original_data.get(col, '')) != str(new_data.get(col, ''))
            for col in self.editable_columns_2 if col not in read_only_fields
        )

        if not has_changes:
            if not messagebox.askyesno("Confirmar sin cambios",
                                    "No se realizaron modificaciones. ¿Desea confirmar igualmente?"):
                window.destroy()
                return

        # Actualizar base de datos
        try:
            cursor = self.conn.cursor()
            set_clause = ", ".join([f"[{col}] = ?" for col in self.editable_invisible_columns_2 if col not in read_only_fields])
            query = f"""
                UPDATE {self.table_name}
                SET {set_clause}, [Estado] = ?
                WHERE Codigo_apertura = ? AND Codigo_cierre = ?
            """
            
            params = [new_data.get(col, None) for col in self.editable_invisible_columns_2 if col not in read_only_fields]
            params.extend([
                'Confirmado',
                str(original_data['Codigo_apertura']),
                str(original_data['Codigo_cierre'])
            ])
            
            cursor.execute(query, params)
            self.conn.commit()

            # Verificar si hubo cambios en Tiempo_horas o Tiempo_minutos
            tiempo_cambiado = (
                str(original_data.get('Tiempo_horas', '')) != str(new_data.get('Tiempo_horas', '')) or
                str(original_data.get('Tiempo_minutos', '')) != str(new_data.get('Tiempo_minutos', ''))
            )

            # Si hubo cambios en tiempo, ejecutar los cálculos de indicadores
            if tiempo_cambiado:
                try:
                    cursor = self.conn.cursor()
                    query_indicadores = f"""
                        UPDATE {self.table_name}
                        SET 
                            -- SAIFI: si conteo_saifi es 0 → todo queda en 0
                            Saifi_contribucion_global = CASE 
                                WHEN conteo_saifi = 0 THEN 0
                                ELSE CAST(Clientes_afectados AS FLOAT) / NULLIF(CAST(Clientes_nacional AS FLOAT), 0)
                            END,

                            Saifi_grupo = CASE 
                                WHEN conteo_saifi = 0 THEN 0
                                ELSE CAST(Clientes_afectados AS FLOAT) / NULLIF(CAST(Clientes_grupo AS FLOAT), 0)
                            END,

                            Saifi_zona = CASE 
                                WHEN conteo_saifi = 0 THEN 0
                                ELSE CAST(Clientes_afectados AS FLOAT) / NULLIF(CAST(Clientes_zona AS FLOAT), 0)
                            END,

                            Saifi_sector = CASE 
                                WHEN conteo_saifi = 0 THEN 0
                                ELSE CAST(Clientes_afectados AS FLOAT) / NULLIF(CAST(Clientes_sector AS FLOAT), 0)
                            END,

                            -- SAIDI: se calculan normal
                            Saidi_contribucion_global = 
                                (CAST(Clientes_afectados AS FLOAT) * CAST(Tiempo_horas AS FLOAT)) 
                                / NULLIF(CAST(Clientes_nacional AS FLOAT), 0),

                            Saidi_grupo = 
                                (CAST(Clientes_afectados AS FLOAT) * CAST(Tiempo_horas AS FLOAT)) 
                                / NULLIF(CAST(Clientes_grupo AS FLOAT), 0),

                            Saidi_zona = 
                                (CAST(Clientes_afectados AS FLOAT) * CAST(Tiempo_horas AS FLOAT)) 
                                / NULLIF(CAST(Clientes_zona AS FLOAT), 0),

                            Saidi_sector = 
                                (CAST(Clientes_afectados AS FLOAT) * CAST(Tiempo_horas AS FLOAT)) 
                                / NULLIF(CAST(Clientes_sector AS FLOAT), 0),
                            cambio_hora = 'SI'

                        WHERE Codigo_apertura = ? AND Codigo_cierre = ?
                    """
                    cursor.execute(query_indicadores, (str(original_data['Codigo_apertura']), str(original_data['Codigo_cierre'])))
                    self.conn.commit()
                    cursor.close()
                except pyodbc.Error as e:
                    self.conn.rollback()
                    messagebox.showerror("Error", f"Error al calcular indicadores: {str(e)}")
            # Actualizar interfaz
            self.revisados()
            window.destroy()
            messagebox.showinfo("Éxito", "Cambios guardados correctamente")

        except pyodbc.Error as e:
            self.conn.rollback()
        
            error_msg = f"Error de base de datos: {str(e)}"
            print(error_msg)
            if ')' in str(e):
                error_msg += "\nPosible error en tipos de datos o formato incorrecto"
            messagebox.showerror("Error", error_msg)
        except Exception as e:
            self.conn.rollback()
            messagebox.showerror("Error", f"Error inesperado: {str(e)}")
        
        


    def mostrar_tabla(self, df):
        self.tree.delete(*self.tree.get_children())
        self.original_data = df.copy()
        try:
            self.estado_col_index = list(df.columns).index('Revisado_operaciones')
        except ValueError:
            self.estado_col_index = None
        date_columns = [col for col in df.columns if 'Fecha' in col]
        for col in date_columns:
            if pd.api.types.is_datetime64_any_dtype(df[col]):
                df[col] = df[col].dt.strftime("%Y-%m-%d %H:%M:%S")

        for _, row in df.iterrows():
            #values = [str(row[col]) for col in self.display_columns]
            values = [cell if cell is not None else "-" for cell in row]
            
            estado = str(row.get('Revisado_operaciones', '')).strip()
            #print(f"Estado fila: '{estado}'")
            if estado == 'Aceptado':
                self.tree.insert("", "end", values=values, tags=("verde",))
            elif estado == 'Rechazado':
                self.tree.insert("", "end", values=values, tags=("rojo",))
            elif estado == 'Pendiente':
                self.tree.insert("", "end", values=values, tags=("naranja",))
            else:
                self.tree.insert("", "end", values=values)

        # Configurar los estilos de cada tag
        self.tree.tag_configure("verde", background="#1c7e02")    # verde claro
        self.tree.tag_configure("rojo", background="#ab0808")     # rojo claro
        self.tree.tag_configure("naranja", background="#333333")  # anaranjado claro

        for col in self.display_columns_2:
            self.tree.heading(col, 
                            text=col.upper(), 
                            command=lambda c=col: self.mostrar_columna_filtrada(c),
                            anchor="center")
        
        #self.auto_adjust_columns()

    # Función para crear una etiqueta con círculo
    def create_label_with_circle(self,container, text, color):
        # Crear el lienzo para el círculo
        canvas = ctk.CTkCanvas(container, width=30, bg = '#333333',height=30, bd=0, highlightthickness=0)
        canvas.create_oval(5, 5, 25, 25, fill=color)  # Dibuja el círculo
        canvas.pack(side="left", padx=5)  # Empaque el lienzo a la izquierda

        # Crear la etiqueta a la derecha del círculo
        label = ctk.CTkLabel(container, text=text, font=("Segoe UI", 12))
        label.pack(side="left", padx=10)  # Empaque la etiqueta a la derecha del círculo


    def abrir_si_existe_sharepoint(self,cod_apertura, cod_cierre, carpeta_intermedia,carpeta_mes):

        carpeta_nombre = f"{cod_apertura}_{cod_cierre}"
        base_path = "/sites/ControlGestin/Documentos compartidos/PRUEBAS"
        current_path = f"{base_path}/{carpeta_intermedia}/{carpeta_mes}/{carpeta_nombre}"

        # Construye la URL de SharePoint (sin verificar existencia) 
        url_base = "https://eneeutcd.sharepoint.com"
        full_url = f"{url_base}{current_path}"
        
        try:
            response = requests.get(full_url, timeout=5)
            content = response.text.lower()
            print(content)

            if response.status_code == 200:
                webbrowser.open(full_url)
            elif response.status_code == 403:
                # Podría existir pero necesitas estar logueado
                webbrowser.open(full_url)
            else:
                messagebox.showinfo("No encontrada", f"La carpeta no está disponible (código {response.status_code}).")
        except requests.RequestException as e:
            messagebox.showerror("Error de red", f"No se pudo verificar la carpeta:\n{e}")

    def color_mouse(self, event):
        item_id = self.tree.identify_row(event.y)
        
        for child in self.tree.get_children():
            estado = str(self.tree.item(child, 'values')[self.estado_col_index]).strip()

            # Define el color base según el estado
            if estado == 'Aceptado':
                base_color = '#1c7e02'
            elif estado == 'Rechazado':
                base_color = '#ab0808'
            elif estado == 'Pendiente':
                base_color = '#0a025d'
            else:
                base_color = '#333333'  # default

            if child == item_id:
                # Color hover personalizado
                self.tree.item(child, tags=('hover_temp',))
                self.tree.tag_configure('hover_temp', background='#3A3A3A')
            else:
                self.tree.item(child, tags=('estado_' + estado.lower(),))
                self.tree.tag_configure('estado_' + estado.lower(), background=base_color)


    def actualizar(self):
        """Actualiza la tabla según la pestaña activa"""
        try:
            self.revisados()    
            messagebox.showinfo("Actualizado", f"Tabla actualizada ", parent=self.root)
            
        except Exception as e:
            self.handle_error(e)

    def show_calendar(self):
        def aplicar_fecha():
            if modo_fecha.get() == "Única":
                fecha = cal_unica.get_date()
                self.filter_fecha_inicio.set(fecha)
                self.filter_fecha_fin.set(fecha)
            else:
                self.filter_fecha_inicio.set(cal_inicio.get_date())
                self.filter_fecha_fin.set(cal_fin.get_date())
            
            top.destroy()
            self.apply_filters()

        top = ctk.CTkToplevel(self.root)
        top.title("Seleccionar Fecha")
        top.geometry("400x380")
        top.configure(bg="#FFFFFF")
        top.resizable(False, False)

        contenedor = ctk.CTkFrame(top, fg_color="#002d6a")
        contenedor.pack(padx=20, pady=20, fill="both", expand=True)

        # Selector de modo
        ctk.CTkLabel(contenedor, text="Seleccione el tipo de filtro:", 
                    font=("Segoe UI", 12)).pack(pady=5)
        
        modo_fecha = ctk.StringVar(value="Única")
        opciones = ["Única", "Rango"]
        
        ctk.CTkSegmentedButton(
            contenedor,
            values=opciones,
            variable=modo_fecha,
            command=lambda v: actualizar_interfaz(),
            fg_color="#E9ECEF",
            selected_color="#007ACC",
            selected_hover_color="#005F9E"
        ).pack(pady=5)

        # Contenedor de calendarios
        calendarios_frame = ctk.CTkFrame(contenedor, fg_color="transparent")
        calendarios_frame.pack(pady=10)

        # Calendario único
        lbl_unica = ctk.CTkLabel(calendarios_frame, text="Seleccione una fecha:")
        cal_unica = DateEntry(
            calendarios_frame,
            date_pattern='yyyy-mm-dd',
            font=("Segoe UI", 12),
            background='white',
            foreground='black'
        )

        # Calendarios de rango (ocultos inicialmente)
        lbl_inicio = ctk.CTkLabel(calendarios_frame, text="Fecha inicial:")
        cal_inicio = DateEntry(
            calendarios_frame,
            date_pattern='yyyy-mm-dd',
            font=("Segoe UI", 12),
            background='white',
            foreground='black'
        )
        
        lbl_fin = ctk.CTkLabel(calendarios_frame, text="Fecha final:")
        cal_fin = DateEntry(
            calendarios_frame,
            date_pattern='yyyy-mm-dd',
            font=("Segoe UI", 12),
            background='white',
            foreground='black'
        )

        def actualizar_interfaz():
            for widget in calendarios_frame.winfo_children():
                widget.pack_forget()
            
            if modo_fecha.get() == "Única":
                lbl_unica.pack(pady=5)
                cal_unica.pack(pady=5)
            else:
                lbl_inicio.pack(pady=5)
                cal_inicio.pack(pady=5)
                lbl_fin.pack(pady=(15,5))
                cal_fin.pack(pady=5)

        actualizar_interfaz()

        # Botones de acción
        botones_frame = ctk.CTkFrame(contenedor, fg_color="transparent")
        botones_frame.pack(pady=10)

        ctk.CTkButton(
            botones_frame,
            text="Cancelar",
            command=top.destroy,
            fg_color="#6C757D",
            hover_color="#5A6268",
            width=100
        ).pack(side="left", padx=10)

        ctk.CTkButton(
            botones_frame,
            text="Aplicar",
            command=aplicar_fecha,
            fg_color="#007ACC",
            hover_color="#005F9E",
            width=100
        ).pack(side="right", padx=10)

        top.transient(self.root)
        top.grab_set()

    def apply_filters(self):
        if self.active_tab.get() == "pendiente":
            self.load_table()
        elif self.active_tab.get() == "Confirmado":
            self.edita_table()
            
    def aplicar_filtro(self):
        self.revisados()

    def clear_filters(self):
        try:
            self.filter_fecha_inicio.set("")
            self.filter_fecha_fin.set("")
            self.filter_codigo.set("")
            self.filter_fecha.set("")
            self.filter_mes.set(self.current_month)
            self.filter_ano.set(self.current_year)
            # Limpiar todos los filtros de columnas
            self.active_filters.clear()
            
            # Cerrar todas las ventanas de filtro de columnas abiertas
            for column in list(self.filter_windows.keys()):
                window = self.filter_windows.pop(column)
                window.destroy()
            
            # Actualizar la tabla aplicando los cambios
            self.apply_filters()
            self.apply_active_filters()  # Asegura aplicar los filtros vacíos
        except Exception as e:
            print(f"Error al limpiar filtros: {e}")
        
    def limpiar_filtros(self):
        self.filter_fecha_inicio.set("")
        self.filter_fecha_fin.set("")
        self.filter_codigo.set("")
        self.filter_fecha.set("")
        self.filter_mes.set(self.current_month) 
        self.filter_ano.set(self.current_year)
        # Limpiar todos los filtros de columnas
        self.active_filters.clear()
        
        # Cerrar todas las ventanas de filtro de columnas abiertas
        for column in list(self.filter_windows.keys()):
            window = self.filter_windows.pop(column)
            window.destroy()
        
        # Actualizar la tabla aplicando los cambios
        self.aplicar_filtro()
        self.aplicar_filtro_activos()  # Asegura aplicar los filtros vacíos
    
    def aplicar_filtro_activos(self):
        # Copia los datos originales para aplicar los filtros
        filtered_df = self.original_data.copy()

        # Aplicar filtros activos en todas las columnas
        for column, values in self.active_filters.items():
            if values:  # Evitar filtros vacíos
                filtered_df = filtered_df[filtered_df[column].astype(str).isin(values)]

        # Actualizar la vista en la tabla (Treeview)
        self.tree.delete(*self.tree.get_children())
        for _, row in filtered_df.iterrows():
            #values = [str(row[col]) for col in self.display_columns]
            values = [cell if cell is not None else "-" for cell in row]
            estado = str(row.get('Revisado_operaciones', '')).strip()
            #print(f"Estado fila: '{estado}'")
            if estado == 'Aceptado':
                self.tree.insert("", "end", values=values, tags=("verde",))
            elif estado == 'Rechazado':
                self.tree.insert("", "end", values=values, tags=("rojo",))
            elif estado == 'Pendiente':
                self.tree.insert("", "end", values=values, tags=("naranja",))
            else:
                self.tree.insert("", "end", values=values)

        # Configurar los estilos de cada tag
        self.tree.tag_configure("verde", background="#1c7e02")    # verde claro
        self.tree.tag_configure("rojo", background="#ab0808")     # rojo claro
        self.tree.tag_configure("naranja", background="#333333")  # anaranjado claro

        # Actualizar el contador de registros
        if self.active_tab_2.get() == "pendiente":
            self.record_count.configure(text=f"Registros: {len(filtered_df)}")
        elif self.active_tab_2.get() == "Confirmado":
            self.record_count.configure(text=f"Registros: {len(filtered_df)}")
        elif self.revision_tab.get() == "si_revision":
            self.record_count.configure(text=f"Registros: {len(filtered_df)}")

        # Resaltar columnas con filtros activos
        for col in self.display_columns:
            if col in self.active_filters:
                self.tree.heading(col, text=f"{col.upper()} ▼", font=('Arial', 10, 'bold'))
            else:
                self.tree.heading(col, text=col.upper(), font=('Arial', 10))

        #Refrescar filtros en todas las columnas para actualizar opciones dinámicamente
        for col in self.active_filters.keys():
            self.mostrar_columna_filtrada(col)
    
    
    def mostrar_columna_filtrada(self, column):
        # Cerrar ventana de filtro si ya está abierta
        if column in self.filter_windows:
            self.filter_windows[column].destroy()

        # Crear ventana de filtro
        filter_win = ctk.CTkToplevel(self.root)
        filter_win.title(f"Filtrar {column}")
        filter_win.geometry("300x400")
        filter_win.transient(self.root)

        # Obtener valores únicos
        unique_values = self.get_unique_values(column)
        if not unique_values:
            ctk.CTkLabel(filter_win, text="No hay valores disponibles").pack(pady=10)
            return

        # 🔎 Entrada de búsqueda
        search_var = ctk.StringVar()
        search_entry = ctk.CTkEntry(filter_win, textvariable=search_var, placeholder_text="Buscar...")
        search_entry.pack(pady=5, padx=5, fill='x')

        # Frame desplazable
        scroll_frame = ctk.CTkScrollableFrame(filter_win)
        scroll_frame.pack(fill='both', expand=True, padx=5, pady=5)

        # Variables y widgets de checkboxes
        check_vars = {}
        checkboxes = {}

        def update_checkboxes():
            search_text = search_var.get().lower()
            for value, checkbox in checkboxes.items():
                if search_text in str(value).lower():
                    checkbox.pack(anchor='w', padx=5, pady=2)
                else:
                    checkbox.pack_forget()

        # Crear checkboxes
        for value in unique_values:
            var = ctk.BooleanVar(value=value in self.active_filters.get(column, []))
            chk = ctk.CTkCheckBox(scroll_frame, text=str(value), variable=var)
            chk.pack(anchor='w', padx=5, pady=2)
            check_vars[value] = var
            checkboxes[value] = chk

        # Escuchar cambios en el texto del Entry
        search_var.trace_add("write", lambda *args: update_checkboxes())

        # Botones de acción
        btn_frame = ctk.CTkFrame(filter_win)
        btn_frame.pack(pady=5, fill='x')

        ctk.CTkButton(btn_frame, text="Aplicar", command=lambda: self.aplicar_filtors_de_columna(column, check_vars, filter_win)).pack(side='left', padx=5)
        ctk.CTkButton(btn_frame, text="Limpiar", command=lambda: self.limpiar_filtors_de_columna(column, filter_win)).pack(side='left', padx=5)

        self.filter_windows[column] = filter_win
    
    def aplicar_filtors_de_columna(self, column, check_vars, window):
        # Obtener valores seleccionados
        selected = [value for value, var in check_vars.items() if var.get()]
        
        if selected:
            self.active_filters[column] = selected
        else:
            self.active_filters.pop(column, None)
        
        window.destroy()
        self.aplicar_filtro_activos()

    def limpiar_filtors_de_columna(self, column, window):
        self.active_filters.pop(column, None)
        window.destroy()
        self.aplicar_filtro_activos()
            
    #Función para actualizar
    def refresh_table(self):
        """Actualiza la tabla según la pestaña activa"""
        try:
            if self.active_tab.get() == "pendiente":
                self.x = 'Pendientes'
                self.load_table()
            elif self.active_tab.get() == "Confirmado":
                self.edita_table()
                self.x = 'Confirmados'
            self.clear_filters()
                
            messagebox.showinfo("Actualizado", f"Tabla de eventos {self.x} actualizados ", parent=self.root)
            
        except Exception as e:
            self.handle_error(e)

    def on_vertical_scroll(self, event):
        self.tree.yview_scroll(-1 * (event.delta // 120), "units")

    def on_horizontal_scroll(self, event):
        self.tree.xview_scroll(-1 * 8*(event.delta // 120), "units")

    def on_hover(self, event):
        item = self.tree.identify_row(event.y)
        self.tree.tag_configure('hover', background='#3A3A3A')
        for child in self.tree.get_children():
            if child == item:
                self.tree.item(child, tags=('hover',))
            else:
                self.tree.item(child, tags=())

    def sort_treeview(self, col):
        data = [(self.tree.set(child, col), child) for child in self.tree.get_children('')]
        
        # Determinar tipo de datos
        try:
            data.sort(key=lambda x: float(x[0]), reverse=self.sort_order)
        except ValueError:
            try:
                data.sort(key=lambda x: datetime.strptime(x[0], "%Y-%m-%d %H:%M:%S"), reverse=self.sort_order)
            except:
                data.sort(reverse=self.sort_order)

        # Reordenar items
        for index, (val, child) in enumerate(data):
            self.tree.move(child, '', index)

        # Cambiar dirección de ordenamiento
        self.sort_order = not self.sort_order

        # Actualizar encabezados
        self.tree.heading(col, 
                        text=col + (" ▼" if self.sort_order else " ▲"),
                        command=lambda: self.sort_treeview(col))

    def on_mousewheel(self, event):
        self.tree.yview_scroll(-1 * (event.delta // 120), "units")

    def on_shift_mousewheel(self, event):
        self.tree.xview_scroll(-1 * (event.delta // 120), "units")

    #Tabla principal de eventos pendientes
    def load_table(self):
        try:
            two_months_ago = datetime.now().replace(day=1).date().strftime('%Y-%m-%d')
            all_columns = self.display_columns + self.hidden_columns

            base_query = f"""
                SELECT {', '.join(f'[{col}]' for col in all_columns)}
                FROM {self.table_name}
                WHERE Estado = 'Pendiente'
                AND Activo = 1
            """

            params = []
            conditions = []

            # Filtro de fecha: si hay un valor en filter_fecha, lo agregamos a las condiciones
            # Filtro de rango de fechas
            if self.filter_fecha_inicio.get() and self.filter_fecha_fin.get():
                conditions.append("CAST(Fecha_apertura AS DATE) BETWEEN ? AND ?")
                params.extend([self.filter_fecha_inicio.get(), self.filter_fecha_fin.get()])

            # Filtro de código: si hay un valor en filter_codigo, lo agregamos a las condiciones
            if self.filter_codigo.get():
                conditions.append("Codigo_apertura LIKE ?")
                params.append(f"%{self.filter_codigo.get()}%")

            # Lista de meses en español
            meses_esp = [
                "Todos los meses", "enero", "febrero", "marzo", "abril", "mayo", "junio", 
                "julio", "agosto", "septiembre", "octubre", "noviembre", "diciembre"
            ]

            # Aquí usas la lista en español para obtener el número del mes
            selected_month = self.filter_mes.get().lower()  # Convertir a minúsculas
            if selected_month != "todos los meses":
                # Convertir el nombre del mes en español a su número correspondiente
                month_number = meses_esp.index(selected_month)  # Esto devuelve el índice del mes
                conditions.append("MONTH(Fecha_apertura) = ?")
                params.append(month_number)
                
            selected_year = self.filter_ano.get()
            if selected_year:
                conditions.append("YEAR(Fecha_apertura) = ?")
                params.append(selected_year)
                
            # Si hay condiciones de filtro, agregarlas a la consulta base
            if conditions:
                base_query += " AND " + " AND ".join(conditions)

            # Ordenar por fecha de apertura y cierre
            base_query += " ORDER BY Fecha_apertura ASC, Fecha_cierre ASC"

            # Ejecutar la consulta y cargar los resultados en un DataFrame
            df = pd.read_sql(base_query, self.conn, params=params)
            
            # Convertir las columnas de fecha
            date_columns = [col for col in df.columns if 'Fecha' in col]
            for col in date_columns:
                df[col] = pd.to_datetime(df[col], errors='coerce')
            
            # Procesar columnas numéricas
            df = df.replace({np.nan: None, '': None,' ':None})
            for col in self.float_columns:
                if col in df.columns:
                    df[col] = df[col].fillna(0).apply(lambda x: round(float(x))).astype(int)
            df['Tiempo_horas'] = df['Tiempo_horas'].apply(lambda x: np.format_float_positional(float(x), precision=2, unique=False, fractional=False, trim='k') if pd.notna(x) else x)
            
            
            
            # Mostrar la tabla
            self.display_table(df)
            self.record_count.configure(text=f"Registros pendientes: {len(df)}")
            return df
        except Exception as e:
            self.handle_error(e)

    #función para exportar
    def export_to_excel(self):
        # Obtener los datos utilizando la función load_data
        # Seleccionar la fuente de datos según el estado activo
        if self.active_tab.get()== 'pendiente':
            data = self.load_table()
        elif self.active_tab.get()== 'Confirmado':
            data = self.edita_table()
        else:
            messagebox.showerror("Error", "Estado desconocido. No se puede exportar.")
            return
        #data = self.load_table()  # Asumiendo que `load_data` retorna un DataFrame
        
        if data is not None:
            # Guardar el DataFrame como un archivo Excel
            file_path = ctk.filedialog.asksaveasfilename(defaultextension=".xlsx",
                                                        filetypes=[("Archivos Excel", "*.xlsx")])
            if file_path:
                try:
                    data.to_excel(file_path, index=False)
                    messagebox.showinfo("Éxito", "Los datos han sido exportados a Excel correctamente.")
                except Exception as e:
                    messagebox.showerror("Error", f"No se pudo exportar el archivo: {e}")
            else:
                messagebox.showwarning("Cancelado", "La exportación ha sido cancelada.")
        else:
            messagebox.showerror("Error", "No se encontraron datos para exportar.")

    #función para exportar
    def export_to_excel_revision(self):
        # Obtener los datos utilizando la función load_data
        # Seleccionar la fuente de datos según el estado activo
        try:
            data = self.revisados()
            #data = self.load_table()  # Asumiendo que `load_data` retorna un DataFrame
            
            if data is not None:
                # Guardar el DataFrame como un archivo Excel
                file_path = ctk.filedialog.asksaveasfilename(defaultextension=".xlsx",
                                                            filetypes=[("Archivos Excel", "*.xlsx")])
                if file_path:
                    try:
                        data.to_excel(file_path, index=False)
                        messagebox.showinfo("Éxito", "Los datos han sido exportados a Excel correctamente.")
                    except Exception as e:
                        messagebox.showerror("Error", f"No se pudo exportar el archivo: {e}")
                else:
                    messagebox.showwarning("Cancelado", "La exportación ha sido cancelada.")
            else:
                messagebox.showerror("Error", "No se encontraron datos para exportar.")
        except Exception as e:
            messagebox.showerror("Error", "Estado desconocido. No se puede exportar.")
            return

    def edita_table(self):
        try:
            two_months_ago = datetime.now().replace(day=1).date().strftime('%Y-%m-%d')
            all_columns = self.display_columns + self.hidden_columns

            base_query = f"""
                SELECT {', '.join(f'[{col}]' for col in all_columns)}
                FROM {self.table_name}
                WHERE Estado = 'Confirmado'
                AND Activo = 1
            """

            params = []
            conditions = []

            # Filtro de fecha
            # Filtro de rango de fechas
            if self.filter_fecha_inicio.get() and self.filter_fecha_fin.get():
                conditions.append("CAST(Fecha_apertura AS DATE) BETWEEN ? AND ?")
                params.extend([self.filter_fecha_inicio.get(), self.filter_fecha_fin.get()])
                
            # Filtro de código
            if self.filter_codigo.get():
                conditions.append("Codigo_apertura LIKE ?")
                params.append(f"%{self.filter_codigo.get()}%")

            # Lista de meses en español
            meses_esp = [
                "Todos los meses", "enero", "febrero", "marzo", "abril", "mayo", "junio", 
                "julio", "agosto", "septiembre", "octubre", "noviembre", "diciembre"
            ]

            # Aquí usas la lista en español para obtener el número del mes
            selected_month = self.filter_mes.get().lower()  # Convertir a minúsculas
            if selected_month != "todos los meses":
                # Convertir el nombre del mes en español a su número correspondiente
                month_number = meses_esp.index(selected_month)  # Esto devuelve el índice del mes
                conditions.append("MONTH(Fecha_apertura) = ?")
                params.append(month_number)
            
            selected_year = self.filter_ano.get()
            if selected_year:
                conditions.append("YEAR(Fecha_apertura) = ?")
                params.append(selected_year)

            # Agregar todas las condiciones de filtro
            if conditions:
                base_query += " AND " + " AND ".join(conditions)

            # Ordenar por fecha de apertura y cierre
            base_query += " ORDER BY Fecha_apertura ASC, Fecha_cierre ASC"

            # Ejecutar la consulta y cargar los resultados en un DataFrame
            df = pd.read_sql(base_query, self.conn, params=params)
            
            # Convertir las columnas de fecha
            date_columns = [col for col in df.columns if 'Fecha' in col]
            for col in date_columns:
                df[col] = pd.to_datetime(df[col], errors='coerce')
            
            df = df.replace({np.nan: None, '': None,' ':None})
            # Procesar columnas numéricas
            for col in self.float_columns:
                if col in df.columns:
                    df[col] = df[col].fillna(0).apply(lambda x: round(float(x))).astype(int)
            df['Tiempo_horas'] = df['Tiempo_horas'].apply(lambda x: np.format_float_positional(float(x), precision=2, unique=False, fractional=False, trim='k') if pd.notna(x) else x)
            
            # Mostrar la tabla
            self.display_table(df)
            self.record_count.configure(text=f"Registros Actualizados: {len(df)}")
            return df
        except Exception as e:
            self.handle_error(e)

    def display_table(self, df):
        self.tree.delete(*self.tree.get_children())
        self.original_data = df.copy()  # Guardar datos originales
        
        # Convertir fechas a string
        date_columns = [col for col in df.columns if 'Fecha' in col]
        for col in date_columns:
            if pd.api.types.is_datetime64_any_dtype(df[col]):
                df[col] = df[col].dt.strftime("%Y-%m-%d %H:%M:%S")
        
        # Insertar datos
        # Diccionario para almacenar registros completos
        self.registros_raw = {}


        # limpiar treeview
        for item in self.tree.get_children():
            self.tree.delete(item)

        for _, row in df.iterrows():

            # 1) Crear id compuesto único
            id_registro = f"{row['Codigo_apertura']}|{row['Codigo_cierre']}"

            # ⛔ 1.5) Ignorar si ya existe en el treeview
            if id_registro in self.tree.get_children():
                continue   # saltar duplicado

            # 2) Guardar todo el registro (aunque no se muestre en el tree)
            self.registros_raw[id_registro] = row.to_dict()

            # 3) Obtener solo las columnas visibles
            values = [str(row[col]) for col in self.display_columns]

            # 4) Insertar en el treeview usando iid = id compuesto
            self.tree.insert(
                "", "end",
                iid=id_registro,        # ← ahora puedes recuperar el registro completo luego
                values=values
            )
        
        # Configurar clic en encabezados para filtros
        for col in self.display_columns:
            self.tree.heading(col, 
                            text=col.upper(), 
                            command=lambda c=col: self.show_column_filter(c),
                            anchor="center")
        
        self.auto_adjust_columns()

    def show_column_filter(self, column):
        # Cerrar ventana de filtro si ya está abierta
        if column in self.filter_windows:
            self.filter_windows[column].destroy()

        # Crear ventana de filtro
        filter_win = ctk.CTkToplevel(self.root)
        filter_win.title(f"Filtrar {column}")
        filter_win.geometry("300x400")
        filter_win.transient(self.root)

        # Obtener valores únicos
        unique_values = self.get_unique_values(column)
        if not unique_values:
            ctk.CTkLabel(filter_win, text="No hay valores disponibles").pack(pady=10)
            return

        # 🔎 Entrada de búsqueda
        search_var = ctk.StringVar()
        search_entry = ctk.CTkEntry(filter_win, textvariable=search_var, placeholder_text="Buscar...")
        search_entry.pack(pady=5, padx=5, fill='x')

        # Frame desplazable
        scroll_frame = ctk.CTkScrollableFrame(filter_win)
        scroll_frame.pack(fill='both', expand=True, padx=5, pady=5)

        # Variables y widgets de checkboxes
        check_vars = {}
        checkboxes = {}

        def update_checkboxes():
            search_text = search_var.get().lower()
            for value, checkbox in checkboxes.items():
                if search_text in str(value).lower():
                    checkbox.pack(anchor='w', padx=5, pady=2)
                else:
                    checkbox.pack_forget()

        # Crear checkboxes
        for value in unique_values:
            var = ctk.BooleanVar(value=value in self.active_filters.get(column, []))
            chk = ctk.CTkCheckBox(scroll_frame, text=str(value), variable=var)
            chk.pack(anchor='w', padx=5, pady=2)
            check_vars[value] = var
            checkboxes[value] = chk

        # Escuchar cambios en el texto del Entry
        search_var.trace_add("write", lambda *args: update_checkboxes())

        # Botones de acción
        btn_frame = ctk.CTkFrame(filter_win)
        btn_frame.pack(pady=5, fill='x')

        ctk.CTkButton(btn_frame, text="Aplicar", command=lambda: self.apply_column_filter(column, check_vars, filter_win)).pack(side='left', padx=5)
        ctk.CTkButton(btn_frame, text="Limpiar", command=lambda: self.clear_column_filter(column, filter_win)).pack(side='left', padx=5)

        self.filter_windows[column] = filter_win


    def get_unique_values(self, column):
        # Aplicar todos los filtros antes de obtener valores únicos
        filtered_df = self.original_data.copy()
        
        for col, values in self.active_filters.items():
            if values:  # Evita filtrar con listas vacías
                filtered_df = filtered_df[filtered_df[col].astype(str).isin(values)]
        
        # Obtener valores únicos de la columna después del filtrado
        return sorted(filtered_df[column].astype(str).unique().tolist(), key=lambda x: x.lower())

    def apply_column_filter(self, column, check_vars, window):
        # Obtener valores seleccionados
        selected = [value for value, var in check_vars.items() if var.get()]
        
        if selected:
            self.active_filters[column] = selected
        else:
            self.active_filters.pop(column, None)
        
        window.destroy()
        self.apply_active_filters()

    def clear_column_filter(self, column, window):
        self.active_filters.pop(column, None)
        window.destroy()
        self.apply_active_filters()
        
    def apply_active_filters(self):
        try:
            # Copia los datos originales para aplicar los filtros
            filtered_df = self.original_data.copy()

            # Aplicar filtros activos en todas las columnas
            for column, values in self.active_filters.items():
                if values:  # Evitar filtros vacíos
                    filtered_df = filtered_df[filtered_df[column].astype(str).isin(values)]

            # Actualizar la vista en la tabla (Treeview)
            self.tree.delete(*self.tree.get_children())
            # for _, row in filtered_df.iterrows():
            #     values = [str(row[col]) for col in self.display_columns]
            #     self.tree.insert("", "end", values=values)
            
            if self.revision_tab.get() == "si_revision":
                self.mostrar_tabla(filtered_df)
            else:
                self.display_table(filtered_df)
            # Actualizar el contador de registros
            if self.active_tab.get() == "pendiente":
                self.record_count.configure(text=f"Registros pendientes: {len(filtered_df)}")
            elif self.active_tab.get() == "Confirmado":
                self.record_count.configure(text=f"Registros Actualizados: {len(filtered_df)}")

            # Resaltar columnas con filtros activos
            # for col in self.display_columns:
            #     heading_text = col.upper()
            #     if col in self.active_filters:
            #         heading_text += " ▼"  # indicar que hay filtro activo
            #     self.tree.heading(col, text=heading_text)

            # Resaltar columnas con filtros activos 
            for col in self.display_columns: 
                if col in self.active_filters: 
                    self.tree.heading(col, text=f"{col.upper()}", font=('Arial', 10, 'bold')) 
                else: 
                    self.tree.heading(col, text=col.upper(), font=('Arial', 10))
            # Refrescar filtros en todas las columnas para actualizar opciones dinámicamente
            for col in self.active_filters.keys():
                self.show_column_filter(col)
        except Exception as e:
            print(f"Error al aplicar filtros activos: {e}")

    def update_filter_options(self, filtered_df):
        """
        Actualiza los valores únicos disponibles en los filtros
        para que reflejen únicamente los datos actualmente visibles.
        """
        self.filtered_values = {
            col: sorted(filtered_df[col].astype(str).unique().tolist(), key=lambda x: x.lower())
            for col in self.display_columns
        }

    def auto_adjust_columns(self):
        for col in self.display_columns:
            max_len = max(
                [len(str(row[col])) for _, row in self.original_data.iterrows()] + [len(col)]
            )
            self.tree.column(col, width=min(max_len * 8, 300))

    def obtener_registros_interrupcion(self):
        """Consulta la base de datos para obtener registros y descripciones."""
        cursor = self.conn.cursor()
        query = "SELECT Registro, Descripcion_falla FROM [GestionControl].[dbo].[CLASIFICACION_INTERRUPCIONES] ORDER BY Registro asc"
        cursor.execute(query)
        registros = cursor.fetchall()
        cursor.close()
        
        # Creamos una estructura que mantenga ambos valores
        return [
            (str(registro), descripcion.strip() if descripcion else "")
            for registro, descripcion in registros
        ]

    def delete_row(self):
        selected = self.tree.selection()
        if not selected:
            messagebox.showwarning("Advertencia", "Seleccione un registro para desactivar")
            return
        
        item = selected[0]
        values = self.tree.item(item)["values"]
        data = dict(zip(self.display_columns, values))
        
        fecha_revision = str(datetime.now().strftime('%Y-%m-%d %H:%M:%S'))
        usuario = f'{self.user} - {fecha_revision}'
        confirm = messagebox.askyesno("Confirmar Eliminación", 
                                    "¿Está seguro de que desea eliminar este registro?\n\n")
        if not confirm:
            return
        
        try:
            cursor = self.conn.cursor()
            query = f"""
                UPDATE {self.table_name}
                SET Activo = 0,
                    Usuario_actualizacion = ?
                WHERE Codigo_apertura = ? AND Codigo_cierre = ?
            """
            cursor.execute(query, (str(usuario),str(data['Codigo_apertura']), str(data['Codigo_cierre'])))
            self.conn.commit()
            
            # Actualizar la interfaz
            if self.tree.exists(item):
                self.tree.delete(item)  # Eliminar de la vista manteniendo los datos en BD
                
            messagebox.showinfo("Éxito", "Registro desactivado correctamente")
            
        except pyodbc.Error as e:
            self.conn.rollback()
            messagebox.showerror("Error", f"Error al desactivar registro: {str(e)}")
        except Exception as e:
            self.conn.rollback()
            messagebox.showerror("Error", f"Error inesperado: {str(e)}")

    def importar_excel_y_subir_sql(self):
        # 1️⃣ Seleccionar archivo Excel
        file_path = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx *.xls")])
        if not file_path:
            messagebox.showwarning("Advertencia", "No se seleccionó ningún archivo.")
            return

        try:
            # 3️⃣ Leer el archivo Excel
            df = pd.read_excel(file_path)
        except Exception:
            messagebox.showerror("Error", "El archivo no es válido o está dañado.")
            return

        # 4️⃣ Encontrar la fila de encabezados
        try:
            fila_header = df[df.iloc[:, 0] == 'OPERADOR DE TURNO'].index[0]
        except IndexError:
            messagebox.showerror("Error", "El archivo no tiene el formato esperado.")
            return

        # 5️⃣ Configurar encabezados y filtrar datos
        df.columns = df.iloc[fila_header]
        print(df.columns)
        df = df[fila_header + 1:].reset_index(drop=True)
        col_equipo_que_opero = df.iloc[:, 1].copy()  # Columna 2 → ahora será Equipo_que_opero
        col_ubicacion = df['EQUIPO QUE OPERO'].copy()  # Columna original
        df['EQUIPO QUE OPERO'] = col_equipo_que_opero  # Sobrescribir con la columna 2
        df['UBICACION'] = col_ubicacion  # Crear nueva columna
        df = df[(df['EVENTO 1'] == 'A') & (df['EVENTO 2'] == 'C')]
        df = df.iloc[:, 8:]

        for i, col in enumerate(df.columns):
            print({col})
            if not isinstance(col, str):
                print(f"Columna sospechosa en posición {i} : {col} (tipo: {type(col)})")
        def limpiar_nombre(col):
            col = unidecode(col).replace("\n", " ").strip().replace(" ", "_").lower()
            return col.capitalize()

        df.columns = [limpiar_nombre(col) for col in df.columns]

        df = df.rename(columns={
            "Grupo": "Grupo_calidad",
            "Tipo": "Tipo_interruptor",
            "Carga_mw": "Carga_MVA",
            "Rele_que_opero": "Relevador",
            "Causa": "Interrupcion",
            "Registro": "Registro_interrupcion",
            "Equipo_que_opero": "Equipo_opero",
            "UBICACION": "Ubicacion"
        })

        df.replace(['--', ''], None, inplace=True)

        df['Estado'] = 'Confirmado'
        df['Origen'] = 'Importado'
        df['Activo'] = 1
        fecha_revision = str(datetime.now().strftime('%Y-%m-%d %H:%M:%S'))
        df['Usuario_actualizacion'] = f'{self.user} - {fecha_revision}'
        df['Codigo_apertura'] = None
        df['Codigo_cierre'] = None
        df.columns = df.columns.str.strip()

        for col in ['Fecha_apertura', 'Fecha_cierre']:
            df[col] = pd.to_datetime(df[col], format='%d/%m/%Y %H:%M:%S', errors='coerce')
            df[col] = df[col].dt.strftime('%Y-%m-%d %H:%M:%S')
            df[col] = df[col].where(df[col].notna(), None)

        df = df[df['Fecha_apertura'] > '2025-04-01 00:00:00']

        df['Clave'] = df.iloc[:, :17].astype(str).agg('-'.join, axis=1)

        def convertir_valores(value):
            if pd.isna(value) or value in ['', None]:
                return None
            try:
                return float(value) if isinstance(value, (int, float, str)) and str(value).replace('.', '', 1).isdigit() else value
            except ValueError:
                return value

        df['Carga_MVA'] = df['Carga_MVA'].apply(convertir_valores)

        conn = self.conn
        cursor = conn.cursor()

        last_num = -1
        last_mes = None
        last_anio = None

        registros_a_insertar = []
        codigos_insertados = []

        # Tomar año y mes actuales del sistema
        hoy = datetime.today()
        año = hoy.year
        mes = hoy.month

        # Calcular primer y último día del mes actual
        fecha_inicio = datetime(año, mes, 1)
        if mes == 12:
            fecha_fin = datetime(año + 1, 1, 1) - timedelta(seconds=1)
        else:
            fecha_fin = datetime(año, mes + 1, 1) - timedelta(seconds=1)

        # Convertir a string para el query
        fecha_inicio_str = fecha_inicio.strftime('%Y-%m-%d %H:%M:%S')
        fecha_fin_str = fecha_fin.strftime('%Y-%m-%d %H:%M:%S')
        
        # Asegúrate de que Fecha_apertura esté en formato datetime
        df['Fecha_apertura'] = pd.to_datetime(df['Fecha_apertura'])

        # Obtener el mes único del DataFrame
        mes_importacion = df['Fecha_apertura'].dt.month.unique()

        if len(mes_importacion) != 1:
            raise ValueError("El archivo contiene datos de más de un mes. Asegúrate de importar un solo mes a la vez.")

        mes_importacion = mes_importacion[0]        

        meses_esp = ["Todos los meses"] + ["Enero", "Febrero", "Marzo", "Abril", "Mayo", "Junio", 
                                      "Julio", "Agosto", "Septiembre", "Octubre", "Noviembre", "Diciembre"]
        self.mes_importado = meses_esp[mes_importacion]  # Mes actual en español

        query_dias = f"""
            SELECT DISTINCT substring(Usuario_actualizacion,1,8) as Usuario_actualizacion, DAY(FECHA_APERTURA) AS DIA, count(*) as CANTIDAD
            FROM {self.table_name}
            WHERE Usuario_actualizacion IS NOT NULL 
            AND MONTH(Fecha_apertura) = {mes_importacion}
            AND Origen = 'Importado' 
            AND Activo = 1
            AND cambio_hora = 'NO'
            AND conteo_saifi = '1'
            group by substring(Usuario_actualizacion,1,8) ,DAY(FECHA_APERTURA)
        """
        df_base = pd.read_sql(query_dias, self.conn)
        print(df_base)

        # Paso 3: construir df_nuevo con tus fechas
        df_nuevo = df[['Fecha_apertura']].copy()
        df_nuevo['DIA'] = df_nuevo['Fecha_apertura'].dt.day
        df_nuevo = df_nuevo.drop_duplicates()
       # df_nuevo['Usuario_actualizacion'] = self.user

        # Paso 4: hacer merge solo por DIA (no por usuario)
        df_duplicados = pd.merge(df_nuevo, df_base, on='DIA')

        # Paso 5: mostrar advertencia si hay coincidencias con otros usuarios
        continuar_importacion = True
           
        if continuar_importacion == True:
            
            # 2️⃣ Crear ventana emergente de carga
            loading_window = Toplevel()
            loading_window.title("Cargando")
            loading_window.geometry("200x100")
            Label(loading_window, text="Cargando datos...").pack(pady=20)
            loading_window.update()

            if df.empty:
                loading_window.destroy()
                messagebox.showinfo("Información", "Esta versión importa datos a partir del 1 de abril de 2025.")
                return
            
            print("✅ Continuando con la importación...")
            # Aquí continuarías con tu lógica de guardado
            for _, row in df.iterrows():
                clave_convertida = str(row['Clave'])
                cursor.execute(f"SELECT COUNT(*) FROM {self.table_name} WHERE Clave  like ?", (clave_convertida,))
                if cursor.fetchone()[0] > 0:
                    print(f"Registro omitido (Clave duplicada): {row['Clave']}")
                    continue

                row["Fecha_apertura"] = pd.to_datetime(row['Fecha_apertura'], errors='coerce')
                fecha_apertura = row['Fecha_apertura']
                anio, mes = fecha_apertura.year, f"{fecha_apertura.month:02d}"

                if last_mes != mes or last_anio != anio:
                    last_num = -1

                cursor.execute(f"""
                    SELECT MAX(Codigo_apertura) FROM {self.table_name}
                    WHERE Codigo_apertura LIKE 'BTE-%-{mes}-{anio}'
                """)
                last_codigo_apertura = cursor.fetchone()
                last_num = int(last_codigo_apertura[0].split('-')[1]) if last_codigo_apertura and last_codigo_apertura[0] else last_num
                last_num += 2

                codigo_apertura = f"BTE-{last_num:04d}-{mes}-{anio}"
                codigo_cierre = f"BTE-{(last_num + 1):04d}-{mes}-{anio}"
                last_num += 2

                row['Codigo_apertura'] = codigo_apertura
                row['Codigo_cierre'] = codigo_cierre
                df.at[_, 'Codigo_apertura'] = codigo_apertura
                df.at[_, 'Codigo_cierre'] = codigo_cierre

                valores = tuple(None if pd.isna(row[col]) or row[col] == '' else row[col] for col in df.columns)
                registros_a_insertar.append(valores)
                codigos_insertados.append(codigo_apertura)

                cursor.execute(f"INSERT INTO {self.table_name} ({', '.join(df.columns)}) VALUES ({', '.join(['?' for _ in df.columns])})", valores)
                conn.commit()

            if codigos_insertados:
                codigos_str = ', '.join(f"'{cod}'" for cod in codigos_insertados)

                update_query_1 = f"""
                    UPDATE {self.table_name}
                    SET Equipo_opero = FORMAT(A.[CODIGO ENERGIS], '0')
                    FROM {self.table_name} BS
                    LEFT JOIN GESTIONCONTROL.DBO.RESTAURADOR A
                    ON BS.Equipo_opero = A.UBICACION AND BS.circuito = A.CIRCUITO
                    WHERE BS.Tipo_interruptor = 'RESTAURADOR'
                    AND ISNUMERIC(BS.Equipo_opero) = 0
                    AND BS.Codigo_apertura IN ({codigos_str})
                """
                #cursor.execute(update_query_1)
                #conn.commit()

                update_query_2 = f"""
                    UPDATE {self.table_name}
                    SET Ubicacion = B.UBICACION
                    FROM {self.table_name} A
                    LEFT JOIN GESTIONCONTROL.DBO.RESTAURADOR B
                    ON A.Equipo_opero = B.[CODIGO ENERGIS] AND A.circuito = B.CIRCUITO
                    WHERE ISNUMERIC(A.Equipo_opero) = 1
                    AND A.Codigo_apertura IN ({codigos_str})
                """
                cursor.execute(update_query_2)
                conn.commit()

                update_query_3 = f"""
                    UPDATE {self.table_name}
                    SET Ubicacion = B.INTERRUPTOR
                    FROM {self.table_name} A
                    LEFT JOIN GESTIONCONTROL.DBO.INTERRUPTOR B
                    ON A.Equipo_opero = B.INTERRUPTOR 
                    WHERE ISNUMERIC(A.Equipo_opero) = 0
                    AND A.Codigo_apertura IN ({codigos_str})
                """
                cursor.execute(update_query_3)
                conn.commit()

                update_query_4 = f"""
                    UPDATE {self.table_name}
                    SET Zona = IIF(ISNUMERIC([Equipo_opero]) = 1, R.Zona, C.ZONA),
                        Sector = IIF(ISNUMERIC([Equipo_opero]) = 1, R.SECTOR,C.SECTOR),
                        Subestacion = IIF(ISNUMERIC([Equipo_opero]) = 1, R.SUBESTACION, C.SUBESTACIÓN)
                    FROM {self.table_name} A
                    LEFT JOIN [GestionControl].[dbo].RESTAURADOR R
                        ON A.[Equipo_opero]=format(R.[CODIGO ENERGIS], '0')
                    LEFT JOIN [GestionControl].[dbo].CIRCUITO C 
                        ON A.CIRCUITO = C.CIRCUITO
                    WHERE A.Codigo_apertura IN ({codigos_str})
                """
                cursor.execute(update_query_4)
                conn.commit()

                # ✅ SOLO AHORA ACTUALIZAMOS CLIENTES (cuando ya `Equipo_opero` está corregido)
                for codigo_apertura in codigos_insertados:
                    codigo_cierre = df[df['Codigo_apertura'] == codigo_apertura]['Codigo_cierre'].values[0]
                    #self.actualizar_clientes_por_codigo(cursor=cursor,codigo_apertura=codigo_apertura,codigo_cierre=codigo_cierre,anio=año,mes=mes,nombre_tabla=self.table_name)

                    # Indicadores también se calculan después
                    #self.actualizar_indicadores_por_codigo(cursor=cursor,codigo_apertura=codigo_apertura,codigo_cierre=codigo_cierre)
                    self.actualizar_lista(cursor=cursor,codigo_apertura=codigo_apertura,codigo_cierre=codigo_cierre)
                    conn.commit()

            loading_window.destroy()
            messagebox.showinfo("Éxito", f"Se insertaron {len(registros_a_insertar)} registros nuevos.")

            cursor.close()
        else:
            messagebox.showerror("Detención", "Elimina los dias que no deseas immportar y vuelve a subir los archivos")
        
    def actualizar_lista(self,cursor,codigo_apertura,codigo_cierre):
        query_lista = f"""
        UPDATE B
        SET B.Interrupcion = A.Descripcion_falla
        FROM {self.table_name} B
        JOIN CLASIFICACION_INTERRUPCIONES A
            ON B.Registro_interrupcion = A.Registro
        WHERE Codigo_apertura = '{codigo_apertura}' AND Codigo_cierre = '{codigo_cierre}'
        """
        cursor.execute(query_lista)
    def actualizar_indicadores_por_codigo(self, cursor, codigo_apertura, codigo_cierre):
        query_indicadores = f"""
        UPDATE {self.table_name}
        SET 
            Saifi_contribucion_global = CAST(Clientes_afectados AS FLOAT) / NULLIF(CAST(Clientes_nacional AS FLOAT), 0),
            Saidi_contribucion_global = (CAST(Clientes_afectados AS FLOAT) * CAST(Tiempo_horas AS FLOAT)) / NULLIF(CAST(Clientes_nacional AS FLOAT), 0),
            Saifi_grupo = CAST(Clientes_afectados AS FLOAT) / NULLIF(CAST(Clientes_grupo AS FLOAT), 0),
            Saidi_grupo = CAST (Clientes_afectados AS FLOAT)*CAST(Tiempo_horas AS FLOAT)/CAST(Clientes_grupo AS FLOAT),
            Saifi_zona = CAST(Clientes_afectados AS FLOAT) / NULLIF(CAST(Clientes_zona AS FLOAT), 0),
            Saidi_zona = CAST (Clientes_afectados AS FLOAT)*CAST(Tiempo_horas AS float)/CAST(Clientes_zona AS FLOAT),
            Saifi_sector = CAST(Clientes_afectados AS FLOAT) / NULLIF(CAST(Clientes_sector AS FLOAT), 0),
            Saidi_sector = (CAST(Clientes_afectados AS FLOAT) * CAST(Tiempo_horas AS FLOAT)) / NULLIF(CAST(Clientes_sector AS FLOAT), 0)
        WHERE Codigo_apertura = '{codigo_apertura}' AND Codigo_cierre = '{codigo_cierre}'
        """
        cursor.execute(query_indicadores)

    def actualizar_clientes_por_codigo(self,cursor, codigo_apertura, codigo_cierre, anio, mes, nombre_tabla):
        query = f"""
        -- Clientes Afectados
        UPDATE {nombre_tabla}
        SET Clientes_afectados = A.Clientes
        FROM {nombre_tabla}
        LEFT JOIN (
            SELECT
                Clientes,
                CONCAT(Año, mes, SUBSTRING(circuito,1,4), SUBSTRING(circuito,6,1), '2L', SUBSTRING(circuito,7,2)) AS Kei
            FROM GESTIONCONTROL.DBO.CLIENTES_CIRCUITO
            WHERE año = {anio}
            UNION ALL
            SELECT
                Clientes,
                CONCAT(Año, mes, Codigo) AS Kei
            FROM GESTIONCONTROL.DBO.CLIENTES_RESTAURADORES
            WHERE año = {anio}
        ) A
        ON TRY_CAST(CONCAT(YEAR(Fecha_apertura), MONTH(Fecha_apertura), Equipo_opero) AS VARCHAR) = TRY_CAST(A.Kei AS VARCHAR)
        WHERE Codigo_apertura = '{codigo_apertura}' AND Codigo_cierre = '{codigo_cierre}';


        -- Clientes por Circuito
        UPDATE {nombre_tabla}
        SET Clientes_circuito = Clientes_afectados
        WHERE Codigo_apertura = '{codigo_apertura}' AND Codigo_cierre = '{codigo_cierre}';

        -- Clientes Sector
        UPDATE {nombre_tabla}
        SET Clientes_sector = CLIENTES_SECTOR.Clientes_sector
        FROM {nombre_tabla}
        LEFT JOIN GESTIONCONTROL.DBO.CLIENTES_SECTOR 
            ON CLIENTES_SECTOR.Sector = {nombre_tabla}.Sector
        WHERE YEAR({nombre_tabla}.Fecha_apertura) = CLIENTES_SECTOR.Año
        AND MONTH({nombre_tabla}.Fecha_apertura) = CLIENTES_SECTOR.Mes
        AND Codigo_apertura = '{codigo_apertura}' AND Codigo_cierre = '{codigo_cierre}';

        -- Clientes Zona
        UPDATE {nombre_tabla}
        SET Clientes_zona = CLIENTES_ZONA.Clientes_zona
        FROM {nombre_tabla}
        LEFT JOIN GESTIONCONTROL.DBO.CLIENTES_ZONA 
            ON CLIENTES_ZONA.ZONA = {nombre_tabla}.ZONA
        WHERE YEAR({nombre_tabla}.Fecha_apertura) = CLIENTES_ZONA.Año
        AND MONTH({nombre_tabla}.Fecha_apertura) = CLIENTES_ZONA.Mes
        AND Codigo_apertura = '{codigo_apertura}' AND Codigo_cierre = '{codigo_cierre}';

        -- Clientes Grupo
        UPDATE {nombre_tabla}
        SET Clientes_grupo = CLIENTES_GRUPO.Clientes_grupo
        FROM {nombre_tabla}
        LEFT JOIN GESTIONCONTROL.DBO.CLIENTES_GRUPO 
            ON CLIENTES_GRUPO.Grupo_calidad = {nombre_tabla}.Grupo_calidad
        WHERE YEAR({nombre_tabla}.Fecha_apertura) = CLIENTES_GRUPO.Año
        AND MONTH({nombre_tabla}.Fecha_apertura) = CLIENTES_GRUPO.Mes
        AND Codigo_apertura = '{codigo_apertura}' AND Codigo_cierre = '{codigo_cierre}';

        -- Clientes Nacional
        UPDATE {nombre_tabla}
        SET Clientes_nacional = CLIENTES_NACIONAL.Clientes_nacional
        FROM {nombre_tabla}
        LEFT JOIN GESTIONCONTROL.DBO.CLIENTES_NACIONAL 
            ON CONCAT(CLIENTES_NACIONAL.Año, RIGHT(CONCAT('0', CLIENTES_NACIONAL.Mes), 2)) = CONCAT(YEAR(Fecha_apertura), RIGHT(CONCAT('0', MONTH(Fecha_apertura)), 2))
        WHERE Codigo_apertura = '{codigo_apertura}' AND Codigo_cierre = '{codigo_cierre}';
        """
        cursor.execute(query)


    def edit_row(self):
        selected = self.tree.selection()
        if not selected:
            messagebox.showwarning("Advertencia", "Seleccione un registro para editar")
            return
        
        iid = selected[0]  # ← ahora este es CódigoApertura|CódigoCierre
        # Validar que exista en registros_raw
        if iid not in self.registros_raw:
            messagebox.showerror("Error", "No se encontraron los datos completos del registro.")
            return
        # Obtener TODOS los datos del registro (incluye columnas no visibles)
        data = self.registros_raw[iid]

        edit_window = ctk.CTkToplevel(self.root)
        edit_window.title("Editar Registro")
        edit_window.geometry("600x700")
        edit_window.grab_set()

        entries = {}
        registros_info = self.obtener_registros_interrupcion()
        
        display_values = [f"{r[0]} - {r[1]}" for r in registros_info]
        registro_mapping = {display: r[0] for display, r in zip(display_values, registros_info)}

        relevador_mapping = {"SIN INTERRUPCIÓN": "SIN INTERRUPCIÓN", "APERTURA REMOTA": "APERTURA REMOTA", "APERTURA LOCAL": "APERTURA LOCAL", "FALLA EN SIN": "FALLA EN SIN", "51ABCN": "51ABCN", "51ABC": "51ABC", "51ABN": "51ABN", "51ACN": "51ACN", "51BCN": "51BCN", "51AN": "51AN", "51BN": "51BN", "51CN": "51CN", "51AB": "51AB", "51AC": "51AC", "51BC": "51BC", "51A": "51A", "51B": "51B", "51C": "51C", "51N": "51N", "51FN": "51FN", "51F": "51F", "21 (DISTANCIA)": "21 (DISTANCIA)", "27 (BAJO VOLTAJE)": "27 (BAJO VOLTAJE)", "50 (SOBRE CORRIENTE INSTANTÁNEO)": "50 (SOBRE CORRIENTE INSTANTÁNEO)", "51 (SOBRE CORRIENTE TEMPORIZADO)": "51 (SOBRE CORRIENTE TEMPORIZADO)", "63 (PRESIÓN)": "63 (PRESIÓN)", "67 (DIRECCIONAL DE SOBRE CORRIENTE)": "67 (DIRECCIONAL DE SOBRE CORRIENTE)", "79 (RECIERRE)": "79 (RECIERRE)", "81 (FRECUENCIA)": "81 (FRECUENCIA)", "86 (BLOQUEO)": "86 (BLOQUEO)", "87 (DIFERENCIAL)": "87 (DIFERENCIAL)", "90 (REGULACIÓN)": "90 (REGULACIÓN)", "NO INDICO": "NO INDICO"}


        relevador_values = [
            "SIN INTERRUPCIÓN", "APERTURA REMOTA", "APERTURA LOCAL", "FALLA EN SIN",
            "51ABCN", "51ABC", "51ABN", "51ACN", "51BCN", "51AN", "51BN", "51CN", 
            "51AB", "51AC", "51BC", "51A", "51B", "51C", "51N", "51FN", "51F", 
            "21 (DISTANCIA)", "27 (BAJO VOLTAJE)", "50 (SOBRE CORRIENTE INSTANTÁNEO)", 
            "51 (SOBRE CORRIENTE TEMPORIZADO)", "63 (PRESIÓN)", "67 (DIRECCIONAL DE SOBRE CORRIENTE)", 
            "79 (RECIERRE)", "81 (FRECUENCIA)", "86 (BLOQUEO)", "87 (DIFERENCIAL)", 
            "90 (REGULACIÓN)", "NO INDICO"
        ]

        class AutoCompleteCombobox(ctk.CTkComboBox):
            def __init__(self, master, values, **kwargs):
                super().__init__(master, values=values, **kwargs)
                self.full_values = values
                self.filtered_values = values
                self.dropdown = None
                self._entry.bind("<KeyRelease>", self.update_filter)
                self._entry.bind("<FocusOut>", lambda e: self.close_dropdown(delay=100))
                self._entry.bind("<Down>", self.open_dropdown)

            def update_filter(self, event=None):
                input_text = self._entry.get().lower()
                self.filtered_values = [v for v in self.full_values if input_text in v.lower()]
                
                if self.dropdown and self.dropdown.winfo_exists():
                    self.update_dropdown()
                else:
                    self.open_dropdown()
                
                self.configure(values=self.filtered_values)

            def open_dropdown(self, event=None):
                if not self.filtered_values:
                    return
                    
                if self.dropdown is None or not self.dropdown.winfo_exists():
                    x = self.winfo_rootx()
                    y = self.winfo_rooty() + self.winfo_height()
                    width = self.winfo_width()
                    
                    self.dropdown = ctk.CTkToplevel(self)
                    self.dropdown.overrideredirect(True)
                    self.dropdown.geometry(f"{width}x200+{x}+{y}")
                    self.dropdown.attributes("-topmost", True)
                    
                    self.scroll_frame = ctk.CTkScrollableFrame(self.dropdown)
                    self.scroll_frame.pack(fill="both", expand=True)
                    self.update_dropdown()
                    
                    self.dropdown.bind("<Button-1>", self.check_click_outside)

            def update_dropdown(self):
                for widget in self.scroll_frame.winfo_children():
                    widget.destroy()
                
                for value in self.filtered_values:
                    btn = ctk.CTkButton(
                        self.scroll_frame,
                        text=value,
                        width=self.winfo_width(),
                        anchor="w",
                        command=lambda v=value: self.select_value(v),
                        fg_color="transparent",
                        hover_color="#3B8ED0",
                        text_color="#FFFFFF",
                        font=("Arial", 12)
                    )
                    btn.pack(pady=1, padx=0)

            def select_value(self, value):
                self.set(value)
                if self.dropdown:
                    self.dropdown.destroy()
                    self.dropdown = None
                self._entry.focus()

            def check_click_outside(self, event):
                if self.dropdown and self.dropdown.winfo_exists():
                    self.dropdown.update_idletasks()
                    d_x = self.dropdown.winfo_x()
                    d_y = self.dropdown.winfo_y()
                    d_width = self.dropdown.winfo_width()
                    d_height = self.dropdown.winfo_height()
                    
                    if not (d_x <= event.x_root <= d_x + d_width and
                            d_y <= event.y_root <= d_y + d_height):
                        self.close_dropdown(delay=0)

            def close_dropdown(self, delay=0):
                if self.dropdown and self.dropdown.winfo_exists():
                    if delay > 0:
                        self.dropdown.after(delay, self.dropdown.destroy)
                    else:
                        self.dropdown.destroy()
                    self.dropdown = None
        
        # Creación de campos de edición
        read_only_fields = ['Circuito', 'Subestacion', 'Ubicacion']
        for col in self.editable_columns:
            frame = ctk.CTkFrame(edit_window, fg_color="transparent")
            frame.pack(pady=8, fill="x", padx=15)
            
            ctk.CTkLabel(frame, text=f"{col}:", width=140, anchor="e").pack(side="left", padx=5)

            if col in read_only_fields:
                value = data.get(col, "")
                label = ctk.CTkLabel(
                    frame,
                    text=value,
                    width=300,
                    fg_color="#343638",
                    corner_radius=3,
                    anchor="w"
                )
                label.pack(side="left")
                continue  # Saltar a la siguiente iteración
            
            # Campo Relevador con lista desplegable
            if col == "Relevador":
                entry = AutoCompleteCombobox(
                    frame, 
                    values=relevador_values, 
                    width=300,
                    dropdown_fg_color="#2B2B2B",
                    button_color="#4A4A4A",
                    border_color="#565B5E",
                    fg_color="#343638",
                    text_color="#FFFFFF"
                )
                
                current_relevador = str(data.get(col, ""))
                #print(f"Valor en data['Relevador']: {current_relevador}")  # Depuración
                current_display = next(
                    (disp for disp, reg in relevador_mapping.items() if reg == current_relevador),
                    ""
                )
                
                entry.set(current_display)
                entry.relevador_mapping = relevador_mapping

                
            # Campo Minutos con botón de cálculo
            elif col == "Tiempo_minutos":
                entry_frame = ctk.CTkFrame(frame, fg_color="transparent")
                entry_frame.pack(side="left")
                
                entry = ctk.CTkEntry(
                    entry_frame,
                    width=250,
                    border_width=1,
                    fg_color="#343638",
                    text_color="#FFFFFF"
                )
                entry.insert(0, str(data.get(col, "")))
                entry.pack(side="left", padx=(0, 5))

                # Capturar la referencia correcta del entry usando lambda
                def crear_calcular_minutos(entry_widget):
                    def calcular_minutos():
                        try:
                            fecha_apertura = datetime.strptime(
                                entries["Fecha_apertura"].get(), 
                                "%Y-%m-%d %H:%M:%S"
                            )
                            fecha_cierre = datetime.strptime(
                                entries["Fecha_cierre"].get(), 
                                "%Y-%m-%d %H:%M:%S"
                            )
                            diferencia = fecha_cierre - fecha_apertura
                            minutos = int(diferencia.total_seconds() / 60+0.5)
                            entry_widget.delete(0, "end")
                            entry_widget.insert(0, str(minutos))
                        except Exception as e:
                            messagebox.showerror(
                                "Error", 
                                f"Error calculando minutos: Verifique los formatos de fecha\nDetalle: {str(e)}"
                            )
                    return calcular_minutos

                ctk.CTkButton(
                    entry_frame,
                    text="🕒 Calcular",
                    width=80,
                    height=28,
                    command=crear_calcular_minutos(entry),  # Pasamos la referencia correcta
                    fg_color="#4A752C",
                    hover_color="#5A8C3F",
                    font=("Arial", 10)
                ).pack(side="left")

            elif col == "Registro_interrupcion":
                entry = AutoCompleteCombobox(
                    frame, 
                    values=display_values, 
                    width=300,
                    dropdown_fg_color="#2B2B2B",
                    button_color="#4A4A4A",
                    border_color="#565B5E",
                    fg_color="#343638",
                    text_color="#FFFFFF"
                )
                current_registro = str(data.get(col, ""))
                current_display = next(
                    (disp for disp, reg in registro_mapping.items() if reg == current_registro),
                    ""
                )
                entry.set(current_display)
                entry.registro_mapping = registro_mapping
                
            elif col == "Observacion":
                entry = ctk.CTkTextbox(
                    frame, 
                    width=300, 
                    height=180, 
                    border_width=1,
                    fg_color="#343638",
                    text_color="#FFFFFF",
                    wrap="word"
                )
                entry.insert("1.0", data.get(col, ""))
                self.agregar_menu_contextual(entry)
                
                def handle_return(event):
                    event.widget.insert("insert", "\n")
                    return "break"
                    
                entry.bind("<Return>", handle_return)
                entry.bind("<KeyRelease-Return>", lambda e: "break")
                
            else:
                entry = ctk.CTkEntry(
                    frame, 
                    width=300, 
                    border_width=1,
                    fg_color="#343638",
                    text_color="#FFFFFF"
                )
                entry.insert(0, str(data.get(col, "")))
            
            entry.pack(side="left")
            entries[col] = entry

        # === CHECKBOX: Cuenta SAIFI o no ===
        check_frame = ctk.CTkFrame(edit_window, fg_color="transparent")
        check_frame.pack(pady=10, fill="x", padx=15)

        ctk.CTkLabel(
            check_frame, 
            text="¿Excluir del SAIFI?:", 
            width=140, 
            anchor="e"
        ).pack(side="left", padx=5)

        # Valor actual en BD (asume 1 = cuenta, 0 = no cuenta)
        valor_actual_saifi = data.get("conteo_saifi")
        print(f'Valor actual SAIFI: {valor_actual_saifi}')
        # Variable CTk
        self.var_saifi = ctk.IntVar(value=valor_actual_saifi)

        # Checkbox (si lo marcan pasa a 0)
        check = ctk.CTkCheckBox(
            check_frame,
            text="Sí",
            variable=self.var_saifi,
            onvalue=0,   # Si lo marca → No cuenta
            offvalue=1,  # Si está desmarcado → Sí cuenta
            checkbox_width=22,
            checkbox_height=22,
            fg_color="#3A3A3A",
            hover_color="#4F4F4F",
            border_color="#AAAAAA",
        )
        check.pack(side="left", padx=10)

        # Añadir a entries para que se guarde correctamente
        entries["conteo_saifi"] = self.var_saifi


        # Botones de acción
        btn_frame = ctk.CTkFrame(edit_window, fg_color="transparent")
        btn_frame.pack(pady=20)
        
        ctk.CTkButton(
            btn_frame,
            text="Guardar Cambios",
            command=lambda: self.save_changes(iid, entries, edit_window),
            width=120,
            height=35,
            fg_color="#28a745",
            hover_color="#218838",
            font=("Arial", 12)
        ).pack(side="left", padx=15)
        
        ctk.CTkButton(
            btn_frame,
            text="Cancelar",
            command=edit_window.destroy,
            width=120,
            height=35,
            fg_color="#dc3545",
            hover_color="#c82333",
            font=("Arial", 12)
        ).pack(side="left", padx=15)


        #edit_window.bind("<Return>", lambda e: self.save_changes(item, entries, edit_window))

    def revisar_ortografia(self, text_widget):
        from spellchecker import SpellChecker

        spell = SpellChecker(language='es')
        text = text_widget.get("1.0", "end-1c")
        words = text.split()

        text_widget.tag_remove("misspelled", "1.0", "end")

        # Detectar palabras mal escritas
        incorrect_words = spell.unknown(words)

        for word in incorrect_words:
            start = "1.0"
            while True:
                start = text_widget.search(word, start, stopindex="end", nocase=True)
                if not start:
                    break
                end = f"{start}+{len(word)}c"
                text_widget.tag_add("misspelled", start, end)
                start = end

        text_widget.tag_config("misspelled", foreground="red", underline=True)

    def save_changes(self, item, entries, window):
        original_values = self.tree.item(item)["values"]
        original_data = dict(zip(self.display_columns, original_values))
        
        # Campos de solo lectura
        read_only_fields = ['Circuito', 'Subestacion', 'Ubicacion']
        all_columns = self.editable_columns + self.hidden_columns
        # Construir new_data excluyendo campos bloqueados
        new_data = {}
        for col in all_columns:
            if col in read_only_fields:
                continue  # Saltar campos no editables
            new_data["conteo_saifi"] = entries["conteo_saifi"].get()
            print(f'Cuenta SAIFI: {new_data["conteo_saifi"] }')
            if col == "Registro_interrupcion":
                # Obtener valor real del registro
                selected_display = entries[col].get().strip()
                registro_mapping = getattr(entries[col], 'registro_mapping', {})
                registro_numero = registro_mapping.get(selected_display, "")
                
                if not registro_numero and selected_display:
                    messagebox.showerror("Error", "Selección inválida para Registro de Interrupción")
                    return
                    
                new_data[col] = registro_numero
                
            elif col == "Observacion":
                new_data[col] = entries[col].get("1.0", "end-1c").strip()
                
            else:
                
                new_data[col] = str(entries[col].get()).strip()
            
        
        fecha_revision = str(datetime.now().strftime('%Y-%m-%d %H:%M:%S'))
        new_data['Usuario_actualizacion'] = f'{self.user} - {fecha_revision}'

        # 1. Validación de tipos de datos
        type_mapping = {
            'int': int,
            'varchar': str,
            'decimal': float,
            'date': lambda x: datetime.strptime(x, '%Y-%m-%d').date(),
            'bit': lambda x: bool(int(x)),
            'datetime': datetime.fromisoformat
        }

        try:
            # Calcular tiempo automático
            if 'Tiempo_minutos' in new_data:
                total_minutos = int(new_data['Tiempo_minutos'])
                horas = total_minutos / 60
                minutos = total_minutos
                new_data['Tiempo_horas'] = horas
                new_data['Tiempo_minutos'] = minutos

            # Obtener clasificación
            if 'Registro_interrupcion' in new_data and new_data['Registro_interrupcion']:
                cursor = self.conn.cursor()
                query = """
                    SELECT [Clasificacion_a], Descripcion_falla
                    FROM [GestionControl].[dbo].[CLASIFICACION_INTERRUPCIONES]
                    WHERE Registro = ?
                """
                cursor.execute(query, (new_data['Registro_interrupcion'],))
                result = cursor.fetchall()
                
                if result:
                    new_data['Clasificacion'] = result[0][0]
                    new_data['Interrupcion'] = result[0][1]

                else:
                    messagebox.showerror('Error', "No se encontró la clasificación para el registro seleccionado")
                    return
                cursor.close()

        except ValueError as e:
            messagebox.showerror("Error", f"Error de conversión de datos: {e}")
            return
        except Exception as e:
            messagebox.showerror("Error", f"Error inesperado: {e}")
            return

        if new_data['Tiempo_horas'] < 0.06:
            new_data['Registro_interrupcion'] = 17
            new_data['Interrupcion'] = 'INSTANTANEA'
            new_data['Clasificacion'] = 'E'
        # Validación de tipos de datos
        type_errors = []
        for col in all_columns:
            if col in read_only_fields:
                continue  # Saltar validación para campos bloqueados    
            sql_type = self.get_sql_type(col).lower()
            base_type = sql_type.split('(')[0].strip()
            converter = type_mapping.get(base_type, str)
            
            try:
                if new_data.get(col, '') == '':
                    new_data[col] = None if base_type != 'varchar' else ''
                else:
                    new_data[col] = converter(new_data[col])
            except Exception as e:
                type_errors.append(f"Columna {col}: Valor '{new_data[col]}' no válido para tipo {sql_type}")

        if type_errors:
            messagebox.showerror("Error de tipos", "\n".join(type_errors))
            return
        print("ORIGINAL DATA:", original_data)
        print("NEW DATA:", new_data)
        # Verificar cambios excluyendo campos bloqueados
        has_changes = any(
            str(original_data.get(col, '')) != str(new_data.get(col, ''))
            for col in all_columns if col not in read_only_fields
        )

        if not has_changes:
            if not messagebox.askyesno("Confirmar sin cambios",
                                    "No se realizaron modificaciones. ¿Desea confirmar igualmente?"):
                window.destroy()
                return

        # Actualizar base de datos
        try:
            cursor = self.conn.cursor()
            set_clause = ", ".join([f"[{col}] = ?" for col in self.editable_invisible_columns if col not in read_only_fields])
            query = f"""
                UPDATE {self.table_name}
                SET {set_clause}, [Estado] = ?
                WHERE Codigo_apertura = ? AND Codigo_cierre = ?
            """
            
            params = [new_data.get(col, None) for col in self.editable_invisible_columns if col not in read_only_fields]
            params.extend([
                'Confirmado',
                str(original_data['Codigo_apertura']),
                str(original_data['Codigo_cierre'])
            ])
            print("QUERY:", query)
            print("PARAMS:", params)
            cursor.execute(query, params)
            self.conn.commit()

            # Verificar si hubo cambios en Tiempo_horas o Tiempo_minutos
            tiempo_cambiado = (
                str(original_data.get('Tiempo_horas', '')) != str(new_data.get('Tiempo_horas', '')) or
                str(original_data.get('Tiempo_minutos', '')) != str(new_data.get('Tiempo_minutos', ''))
            )

            # Si hubo cambios en tiempo, ejecutar los cálculos de indicadores
            if tiempo_cambiado:
                try:
                    cursor = self.conn.cursor()
                    query_indicadores = f"""
                        UPDATE {self.table_name}
                        SET 
                            -- SAIFI: si conteo_saifi es 0 → todo queda en 0
                            Saifi_contribucion_global = CASE 
                                WHEN conteo_saifi = 0 THEN 0
                                ELSE CAST(Clientes_afectados AS FLOAT) / NULLIF(CAST(Clientes_nacional AS FLOAT), 0)
                            END,

                            Saifi_grupo = CASE 
                                WHEN conteo_saifi = 0 THEN 0
                                ELSE CAST(Clientes_afectados AS FLOAT) / NULLIF(CAST(Clientes_grupo AS FLOAT), 0)
                            END,

                            Saifi_zona = CASE 
                                WHEN conteo_saifi = 0 THEN 0
                                ELSE CAST(Clientes_afectados AS FLOAT) / NULLIF(CAST(Clientes_zona AS FLOAT), 0)
                            END,

                            Saifi_sector = CASE 
                                WHEN conteo_saifi = 0 THEN 0
                                ELSE CAST(Clientes_afectados AS FLOAT) / NULLIF(CAST(Clientes_sector AS FLOAT), 0)
                            END,

                            -- SAIDI: se calculan normal
                            Saidi_contribucion_global = 
                                (CAST(Clientes_afectados AS FLOAT) * CAST(Tiempo_horas AS FLOAT)) 
                                / NULLIF(CAST(Clientes_nacional AS FLOAT), 0),

                            Saidi_grupo = 
                                (CAST(Clientes_afectados AS FLOAT) * CAST(Tiempo_horas AS FLOAT)) 
                                / NULLIF(CAST(Clientes_grupo AS FLOAT), 0),

                            Saidi_zona = 
                                (CAST(Clientes_afectados AS FLOAT) * CAST(Tiempo_horas AS FLOAT)) 
                                / NULLIF(CAST(Clientes_zona AS FLOAT), 0),

                            Saidi_sector = 
                                (CAST(Clientes_afectados AS FLOAT) * CAST(Tiempo_horas AS FLOAT)) 
                                / NULLIF(CAST(Clientes_sector AS FLOAT), 0),
                            cambio_hora = 'SI'

                        WHERE Codigo_apertura = ? AND Codigo_cierre = ?
                    """
                    cursor.execute(query_indicadores, (str(original_data['Codigo_apertura']), str(original_data['Codigo_cierre'])))
                    self.conn.commit()
                    cursor.close()
                except pyodbc.Error as e:
                    self.conn.rollback()
                    messagebox.showerror("Error", f"Error al calcular indicadores: {str(e)}")
            # Actualizar interfaz
            if original_data['Estado'] == 'Confirmado':
                self.edita_table()
            else:
                self.load_table()
                
            window.destroy()
            messagebox.showinfo("Éxito", "Cambios guardados correctamente")

        except pyodbc.Error as e:
            self.conn.rollback()
            error_msg = f"Error de base de datos: {str(e)}"
            if ')' in str(e):
                error_msg += "\nPosible error en tipos de datos o formato incorrecto"
            messagebox.showerror("Error", error_msg)
        except Exception as e:
            self.conn.rollback()
            messagebox.showerror("Error", f"Error inesperado: {str(e)}")
        
    def get_sql_type(self, column_name):
        """
        Obtener el tipo de dato SQL para una columna
        Debes implementar esta función según tu estructura de base de datos
        Ejemplo usando información del schema:
        """
        query = f"""
            SELECT DATA_TYPE 
            FROM INFORMATION_SCHEMA.COLUMNS
            WHERE TABLE_NAME = '{self.table_name}' 
            AND COLUMN_NAME = '{column_name}'
        """
        cursor = self.conn.cursor()
        cursor.execute(query)
        return cursor.fetchone()[0]

    def copy_row_to_clipboard(self, event=None):
        # Obtener fila seleccionada
        selected = self.tree.selection()
        if not selected:
            return
        
        # Obtener todos los valores de la fila
        item = selected[0]
        values = self.tree.item(item, "values")
        
        # Convertir a texto separado por tabs
        row_text = "\t".join(str(value) for value in values)
        
        # Copiar al portapapeles
        self.root.clipboard_clear()
        self.root.clipboard_append(row_text)
        
        # Mantener el contenido en el portapapeles después de cerrar la app
        self.root.update()

    def handle_error(self, error):
        error_msg = re.sub(r"\([^)]*\)", "", str(error)).strip()
        messagebox.showerror("Error", f"Error: {error_msg}")

if __name__ == "__main__":
    root = ctk.CTk()
    app = SQLApp(root)
    root.mainloop()