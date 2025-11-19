import pyodbc
import calendar
import customtkinter as ctk
from tkinter import messagebox, ttk
import pandas as pd
import re
import os
import requests
import base64
import webbrowser
from datetime import datetime, timedelta
from unidecode import unidecode
import numpy as np
from tkinter import filedialog
from tkinter import filedialog, Toplevel,Label
from tkcalendar import DateEntry
from dateutil.relativedelta import relativedelta
from datetime import datetime, timedelta
import tkinter as tk
from datetime import datetime
import locale
from tkinter import simpledialog, messagebox, Toplevel, Label, Button
from office365.runtime.auth.authentication_context import AuthenticationContext
from office365.sharepoint.client_context import ClientContext
from office365.sharepoint.client_context import ClientContext
from office365.runtime.auth.authentication_context import AuthenticationContext
from office365.sharepoint.files.file import File


ctk.set_appearance_mode("Dark")
ctk.set_default_color_theme("dark-blue")

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
        #self.root.attributes('-fullscreen', True)
        # Configuración inicial
        self.table_name = "BITACORA_SCADA"
        self.editable_columns = ["Fecha_apertura",'Fecha_cierre',"Tiempo_minutos",'Circuito', 'Subestacion','Ubicacion',"Carga_MVA", "Registro_interrupcion",'Relevador','Observacion','Comentario_sector','Revisado_operaciones','Comentario_operaciones']
        self.editable_invisible_columns =['Fecha_apertura','Fecha_cierre',"Tiempo_horas","Tiempo_minutos", "Carga_MVA",'Registro_interrupcion',"Clasificacion",'Relevador','Interrupcion','Observacion','Usuario_actualizacion']
        self.display_columns = [
            'Codigo_apertura', 'Codigo_cierre', 'Fecha_apertura', 'Fecha_cierre',
            'Tiempo_horas', 'Tiempo_minutos', 'Sector', 'Subestacion',
            'Circuito', 'Tipo_interruptor', 'Equipo_opero','Ubicacion',
            'Carga_MVA', 'Relevador', 'Interrupcion', 'Clasificacion',
            'Registro_interrupcion', 'Observacion','Revisado_operaciones','Comentario_operaciones','Comentario_sector','Revision','Revisado_sector'
        ]
        self.float_columns = ['Tiempo_minutos','Grupo_calidad']
        self.conn = None
        
        self.sort_column = None
        self.sort_order = False  # False = ascendente, True = descendente
        self.soporte = 'Si'
        self.pendiente = 'Pendiente'

        self.filter_windows = {}  # Para mantener ventanas de filtro abiertas
        self.active_filters = {}  # Diccionario de filtros activos {columna: valores}
        self.original_data = pd.DataFrame()  # Copia de los datos sin filtrar
        self.sharepoint_user = ''
        self.sharepoint_url = "https://eneeutcd.sharepoint.com/sites/ControlGestin"
        self.sharepoint_folders = []
        self.site_url = "/sites/ControlGestin"
        self.folder_url = "/Documentos compartidos/PRUEBAS" 
        self.share = 'No'
        self.create_login_interface()

    def create_login_interface(self):
        self.login_frame = ctk.CTkFrame(self.root, fg_color="#2B2B2B")
        self.login_frame.pack(fill="both", expand=True, padx=20, pady=20)

        ctk.CTkLabel(self.login_frame, 
                    text="Conexión SQL Server", 
                    font=("Arial", 20, "bold")).pack(pady=20)

        # Valores fijos
        server_value = "192.168.100.7"
        database_value = "GestionControl"

        # Mostrar valores fijos
        fixed_info = [
            ("Servidor", server_value),
            ("Base de Datos", database_value)
        ]
        for text, value in fixed_info:
            frame = ctk.CTkFrame(self.login_frame, fg_color="transparent")
            frame.pack(pady=5)
            
            ctk.CTkLabel(frame, 
                        text=text + ":", 
                        width=140, 
                        anchor="e").pack(side="left", padx=5)
            
            ctk.CTkLabel(frame, 
                        text=value, 
                        width=250, 
                        fg_color="#3E3E3E", 
                        corner_radius=5, 
                        anchor="w").pack(side="left", padx=5)

        # Campos editables
        self.entries = {}
        user_pass_inputs = [
            ("Usuario", "user_entry"),
            ("Contraseña", "pass_entry")
        ]
        
        for text, name in user_pass_inputs:
            frame = ctk.CTkFrame(self.login_frame, fg_color="transparent")
            frame.pack(pady=5)
            
            ctk.CTkLabel(frame, 
                        text=text + ":", 
                        width=140, 
                        anchor="e").pack(side="left", padx=5)
            
            entry = ctk.CTkEntry(frame, width=250, show="*" if name == "pass_entry" else "")
            entry.pack(side="left")
            # Bindear Enter key a validate_and_connect
            entry.bind("<Return>", lambda event: self.validate_and_connect()) 
            
            self.entries[name] = entry
            
            # Botón para mostrar/ocultar contraseña
            if name == "pass_entry":
                self.show_password = False
                
                def toggle_password():
                    self.show_password = not self.show_password
                    entry.configure(show="" if self.show_password else "*")
                    toggle_btn.configure(text="🔓" if self.show_password else "🔒")  # Cambia el ícono
                
                toggle_btn = ctk.CTkButton(frame, text="🔒", width=40, command=toggle_password)
                toggle_btn.pack(side="left", padx=5)
        
        btn_frame = ctk.CTkFrame(self.login_frame, fg_color="transparent")
        btn_frame.pack(pady=20)
        
        ctk.CTkButton(btn_frame, 
                    text="Conectar", 
                    command=self.validate_and_connect,
                    fg_color="#2E8B57",
                    hover_color="#245c3d",
                    corner_radius=8,
                    font=("Arial", 12, "bold")).pack(side="left", padx=10)

    def validate_and_connect(self):
        self.user = self.entries['user_entry'].get().strip()
        password = self.entries['pass_entry'].get().strip()
        
        if not self.user or not password:
            messagebox.showerror("Error de conexión", "Por favor, complete todos los campos.")
            return
        
        self.connect_to_sql()

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
        
        # Mostrar la fecha seleccionada en un Label
        #self.fecha_label = ctk.CTkLabel(
        #    filter_frame, 
        #    textvariable=self.filter_fecha, 
        #    width=120
        #)
        #self.fecha_label.pack(side="left", padx=5)
        
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
            command=lambda _: self.apply_filters() 
        )
        self.filter_mes.set(self.current_month)  # Establecer el mes actual como valor por defecto
        mes_menu.pack(side="left", padx=5)
        
        # Botón Limpiar Filtros
        ctk.CTkButton(
            filter_frame,
            text="Limpiar Filtros",
            command=self.clear_filters,
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
        self.active_tab = ctk.StringVar(value="pendiente")
        
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
            self.button_container,
            text="🔄 Registros no revisados",
            command=lambda: [self.active_tab.set("pendiente"),self.load_table(), update_button_style(pendientes_btn)],
            fg_color="#007ACC",
            hover_color="#005F9E",
            corner_radius=6,
            font=("Segoe UI Semibold", 12),
            width=160,
            height=32
        )
        pendientes_btn.pack(side="left", padx=5)

        confirmados_btn = ctk.CTkButton(
            self.button_container,
            text="✅ Registros revisados",
            command=lambda: [self.active_tab.set("Confirmado"),self.edita_table(), update_button_style(confirmados_btn)],
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
            text="Verificar/Cambiar",
            command=self.edit_row,
            fg_color="#4b2f4d",
            hover_color="#29760a",
            corner_radius=6,
            font=("Segoe UI Semibold", 12),
            width=160,
            height=32
        ).pack(side="left", padx=5)
    
        #ctk.CTkButton(
        ##    self.button_container,
        ##    command=self.revisar_multiple_rows,
        #    fg_color="#037C60",
        #    hover_color="#29760a",
        #    corner_radius=6,
        #    font=("Segoe UI Semibold", 12),
        #    width=180,
        #    height=32
        #).pack(side="left", padx=10)

        # Añadir botón de actualización
        ctk.CTkButton(
            self.button_container,
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
            self.button_container,
            text="📥 Exportar a Excel",
            command=self.export_to_excel,
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
                                selectmode="extended")
        
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
            
    def clear_filters(self):
        self.filter_fecha_inicio.set("")
        self.filter_fecha_fin.set("")
        self.filter_codigo.set("")
        self.filter_fecha.set("")
        self.filter_mes.set(self.current_month) 
        # Limpiar todos los filtros de columnas
        self.active_filters.clear()
        
        # Cerrar todas las ventanas de filtro de columnas abiertas
        for column in list(self.filter_windows.keys()):
            window = self.filter_windows.pop(column)
            window.destroy()
        
        # Actualizar la tabla aplicando los cambios
        self.apply_filters()
        self.apply_active_filters()  # Asegura aplicar los filtros vacíos
    
    def refresh_table(self):
        """Actualiza la tabla según la pestaña activa"""
        try:
            if self.active_tab.get() == "pendiente":
                self.x = 'Pendientes'
                self.load_table()
            elif self.active_tab.get() == "Confirmado":
                self.edita_table()
                self.x = 'Confirmados'
                
            messagebox.showinfo("Actualizado", f"Tabla de eventos {self.x} actualizados ", parent=self.root)
            
        except Exception as e:
            self.handle_error(e)

    def on_vertical_scroll(self, event):
        self.tree.yview_scroll(-1 * (event.delta // 120), "units")

    def on_horizontal_scroll(self, event):
        self.tree.xview_scroll(-1 * 8*(event.delta // 120), "units")

    def on_hover(self, event):
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

    def load_table(self):
        # Elimina los círculos/etiquetas si existen
        for widget in self.button_container.winfo_children():
            if isinstance(widget, (ctk.CTkCanvas, ctk.CTkLabel)):
                widget.destroy()
        try:
            two_months_ago = datetime.now().replace(day=1).date().strftime('%Y-%m-%d')
            base_query = f"""
                SELECT {', '.join(f'[{col}]' for col in self.display_columns)} 
                FROM {self.table_name}
                WHERE Estado = 'Confirmado' and Revisado_sector is null
                AND Activo = 1 and Registro_interrupcion not in (18,19,36,37,16,17,40,45) and Clasificacion != 'P'
            """

            params = []
            conditions = []

            cursor = self.conn.cursor()
            query_2 = f"""
                SELECT Sector
                FROM Usuarios_sectores
                WHERE Usuario = ? order by ID 
            """
            cursor.execute(query_2, (self.user,))
            resultados = cursor.fetchall()
            self.conn.commit()
    
            if not resultados:
                messagebox.showerror("Error", f"El usuario '{self.user}' no tiene permiso, o no está registrado en ningún sector.")
                return
            else:
                sectores = [r[0] for r in resultados]
                print(sectores)
                self.sectores_ = ['DANLI-TEGUCIGALPA','CHOLUTECA','COMAYAGUA',
                    'JUTICALPA','TOCOA-LA CEIBA','SAN PEDRO SULA',
                    'VILLANUEVA','SANTA ROSA DE COPAN',
                    'SANTA CRUZ DE YOJOA-EL PROGRESO']
            if len(sectores)>1:
                    self.carpetas = f'{sectores[0].upper()}-{sectores[1].upper()}'
            if sectores != ['TODOS']:
                self.carpetas = f'{sectores[0].upper()}'
            self.todos = 0
            if sectores == ['TODOS']:
                self.todos = 1
             
            # Filtro de fecha
            if self.filter_fecha.get():
                conditions.append("CAST(Fecha_apertura AS DATE) = ?")
                params.append(self.filter_fecha.get())
                
            # Filtro de rango de fechas
            if self.filter_fecha_inicio.get() and self.filter_fecha_fin.get():
                conditions.append("CAST(Fecha_apertura AS DATE) BETWEEN ? AND ?")
                params.extend([self.filter_fecha_inicio.get(), self.filter_fecha_fin.get()])

            # Filtro de código
            if self.filter_codigo.get():
                conditions.append("Codigo_apertura LIKE ?")
                params.append(f"%{self.filter_codigo.get()}%")

            # Filtro por sector
            if 'TODOS' not in sectores:
                placeholders = ','.join(['?'] * len(sectores))
                conditions.append(f"Sector IN ({placeholders})")
                params.extend(sectores)
            
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
            self.record_count.configure(text=f"Registros: {len(df)}")
            return df
        except Exception as e:
            self.handle_error(e)

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

    # Función para crear una etiqueta con círculo
    def create_label_with_circle(self,container, text, color):
        # Crear el lienzo para el círculo
        canvas = ctk.CTkCanvas(container, width=30, bg = '#333333',height=30, bd=0, highlightthickness=0)
        canvas.create_oval(5, 5, 25, 25, fill=color)  # Dibuja el círculo
        canvas.pack(side="left", padx=5)  # Empaque el lienzo a la izquierda

        # Crear la etiqueta a la derecha del círculo
        label = ctk.CTkLabel(container, text=text, font=("Segoe UI", 12))
        label.pack(side="left", padx=10)  # Empaque la etiqueta a la derecha del círculo


    def revisar_multiple_rows(self):
        seleccionados = self.tree.selection()
        print(seleccionados)
        if not seleccionados:
            messagebox.showwarning("Advertencia", "Seleccione al menos un registro para revisar.")
            return

        for item in seleccionados:
            valores = self.tree.item(item)["values"]
            print(valores)
            data = dict(zip(self.display_columns, valores))

            self.revisado_valor = f"{self.user} - revisado"
            self.revisado = 'revisado'
            self.pendiente = None
            #self.revisado_valor = self.ocupa_evidencia()
            fecha_revision = str(datetime.now().strftime('%Y-%m-%d %H:%M:%S'))
            self.revisado_valor = f"{self.revisado_valor} - {fecha_revision}"
            try:
                cursor = self.conn.cursor()
                query = f"""
                    UPDATE {self.table_name}
                    SET Revisado_sector = ?, Revisado_operaciones = ?,Revision = ?
                    WHERE Codigo_apertura = ? AND Codigo_cierre = ?
                """
                print('anteeeeeeeeeeeeeeees')
                cursor.execute(query, (
                    self.revisado_valor,
                    self.pendiente,
                    self.revisado,
                    str(data['Codigo_apertura']),
                    str(data['Codigo_cierre'])
                ))
                self.conn.commit()
                print('holaaaaaaaaaaaaaaaaaaaaaaaaaa')
                if self.tree.exists(item) and self.active_tab.get() != "Confirmado":
                    self.tree.delete(item)
                print('aqui entre')
            except Exception as e:
                self.conn.rollback()
                messagebox.showerror("Error", f"Error al revisar registro: {e}")
        messagebox.showinfo("Éxito", f"{len(seleccionados)} registros fueron marcados como revisados.")

        if self.active_tab.get()== 'pendiente':
            data = self.load_table()
        elif self.active_tab.get()== 'Confirmado':
            data = self.edita_table()


    def edita_table(self):        
        # Verifica si las etiquetas con círculos ya existen en el contenedor
        # Elimina cualquier etiqueta y canvas previos en el contenedor
        for widget in self.button_container.winfo_children():
            if isinstance(widget, (ctk.CTkCanvas, ctk.CTkLabel)):
                widget.destroy()

        # Crea nuevas etiquetas con círculo
        self.create_label_with_circle(self.button_container, "Pendiente", "#0a025d")  # Naranja
        self.create_label_with_circle(self.button_container, "Aceptado", "#1c7e02")  # Verde
        self.create_label_with_circle(self.button_container, "Rechazado", "#ab0808")  # Rojo
    
        try:
            two_months_ago = datetime.now().replace(day=1).date().strftime('%Y-%m-%d')
            base_query = f"""
                SELECT {', '.join(f'[{col}]' for col in self.display_columns)} 
                FROM {self.table_name}
                WHERE Estado = 'Confirmado' and Revisado_sector is not null
                AND Activo = 1
            """

            params = []
            conditions = []

            cursor = self.conn.cursor()
            query_2 = f"""
                SELECT Sector
                FROM Usuarios_sectores
                WHERE Usuario = '{self.user}' order by ID
            """
            cursor.execute(query_2)
            resultado = cursor.fetchall()
            print(f'sectores {resultado}')
            self.conn.commit()

            if resultado is None:
                messagebox.showerror("Error", f"El usuario '{self.user}' no Tiene permiso, o no está registrado en el Sector.")
                return  # o manejar como sea apropiado
            else:
                sector = resultado[0]
                sectores = [r[0] for r in resultado]
                print(f'sectores {resultado}')
                print(f'sectores {sectores}')
            
            # Filtro de fecha: si hay un valor en filter_fecha, lo agregamos a las condiciones
            if self.filter_fecha.get():
                conditions.append("CAST(Fecha_apertura AS DATE) = ?")
                params.append(self.filter_fecha.get())

            # Filtro de rango de fechas
            if self.filter_fecha_inicio.get() and self.filter_fecha_fin.get():
                conditions.append("CAST(Fecha_apertura AS DATE) BETWEEN ? AND ?")
                params.extend([self.filter_fecha_inicio.get(), self.filter_fecha_fin.get()])

            # Filtro de código: si hay un valor en filter_codigo, lo agregamos a las condiciones
            if self.filter_codigo.get():
                conditions.append("Codigo_apertura LIKE ?")
                params.append(f"%{self.filter_codigo.get()}%")
            
            # Filtro por sector
            if 'TODOS' not in sectores:
                placeholders = ','.join(['?'] * len(sectores))
                print(placeholders)
                conditions.append(f"Sector IN ({placeholders})")
                params.extend(sectores)
            
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

            # Si hay condiciones de filtro, agregarlas a la consulta base
            if conditions:
                base_query += " AND " + " AND ".join(conditions)

            # Ordenar por fecha de apertura y cierre
            base_query += " ORDER BY CAST(SUBSTRING(Revisado_sector, CHARINDEX('-', Revisado_sector, CHARINDEX('-', Revisado_sector) + 1) + 2, 19) AS DATEtime) asc"
            
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
            self.record_count.configure(text=f"Registros: {len(df)}")
            return df
        except Exception as e:
            print(f'{e}')
            self.handle_error(e)

    def display_table(self, df):
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
        if self.active_tab.get() == "pendiente":
            self.record_count.configure(text=f"Registros: {len(filtered_df)}")
        elif self.active_tab.get() == "Confirmado":
            self.record_count.configure(text=f"Registros Actualizados: {len(filtered_df)}")

        # Resaltar columnas con filtros activos
        for col in self.display_columns:
            if col in self.active_filters:
                self.tree.heading(col, text=f"{col.upper()} ▼", font=('Arial', 10, 'bold'))
            else:
                self.tree.heading(col, text=col.upper(), font=('Arial', 10))

        # Refrescar filtros en todas las columnas para actualizar opciones dinámicamente
        for col in self.active_filters.keys():
            self.show_column_filter(col)

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

    def ocupa_evidencia(self):
        # Crear ventana personalizada
        respuesta = {"accion": None}

        def confirmar():
            respuesta["accion"] = "revisado"
            ventana.destroy()

        def solicitar_cambio():
            respuesta["accion"] = "cambio"
            ventana.destroy()

        ventana = Toplevel(self.root)
        ventana.title("Revisión")
        ventana.geometry("350x120")
        ventana.grab_set()  # Bloquea interacción con ventana principal

        Label(ventana, text="¿Es una exclusión que necesita Soporte?").pack(pady=10)

        btn_frame = tk.Frame(ventana)
        btn_frame.pack(pady=5)

        Button(btn_frame, text="Si", width=10, command=confirmar).pack(side="left", padx=10)
        Button(btn_frame, text="No", width=10, command=solicitar_cambio).pack(side="right", padx=10)

        self.root.wait_window(ventana)  # Espera hasta que se cierre la ventana

        if respuesta["accion"] is None:
            return  # El usuario cerró la ventana sin hacer nada

        if respuesta["accion"] == "revisado":
            self.revisado_valor = f"{self.user} - revisado, adjunta soporte "
            self.pendiente = None
            self.create_login_sharepoint_interface() 
        else:
            self.revisado_valor = f"{self.user} - revisado"
            self.edit_window.destroy()
            self.pendiente = None
            messagebox.showinfo("Éxito", "Cambios guardados correctamente")
        #ventana.grab_set()             # Bloquea la interacción con la ventana principal
        #self.root.wait_window(ventana) # Espera a que se cierre `top`
        return self.revisado_valor
    
    def soporte_comentario(self):
        # Crear ventana personalizada
        respuesta = {"accion": None}

        def confirmar():
            respuesta["accion"] = "revisado"
            ventana.destroy()

        def solicitar_cambio():
            respuesta["accion"] = "cambio"
            ventana.destroy()

        ventana = Toplevel(self.root)
        ventana.title("Revisión")
        ventana.geometry("350x120")
        ventana.grab_set()  # Bloquea interacción con ventana principal

        Label(ventana, text="¿Necesita Soporte o solo un comentario?").pack(pady=10)

        btn_frame = tk.Frame(ventana)
        btn_frame.pack(pady=5)

        Button(btn_frame, text="Comentario", width=10, command=confirmar).pack(side="left", padx=10)
        Button(btn_frame, text="Soporte y comentario", width=15, command=solicitar_cambio).pack(side="right", padx=10)

        self.root.wait_window(ventana)  # Espera hasta que se cierre la ventana

        if respuesta["accion"] is None:
            return
        if respuesta["accion"] == "revisado":
            self.soporte = 'No'
            self.pendiente = 'Pendiente'
            self.revisado_valor = f"{self.user} - solicitó cambio sin soporte"
            self.revisado = 'Esclarecer'
            fecha_revision = str(datetime.now().strftime('%Y-%m-%d %H:%M:%S'))
            self.revisado_valor = f"{self.revisado_valor} - {fecha_revision}"
            self.create_update_data_interface()
            self.edit_window.destroy()
        else:
            self.soporte = 'Si'
            self.pendiente = 'Pendiente'
            self.revisado_valor = f"{self.user} - solicitó cambio con soporte"
            self.revisado = 'Exclusión'
            fecha_revision = str(datetime.now().strftime('%Y-%m-%d %H:%M:%S'))
            self.revisado_valor = f"{self.revisado_valor} - {fecha_revision}"
            self.create_update_data_interface()
            self.edit_window.destroy()
            return self.revisado_valor
            


    def revisar_row(self):
        selected = self.tree.selection()
        if not selected:
            messagebox.showwarning("Advertencia", "Seleccione un registro para revisar")
            return
        self.soporte == 'Si'
        self.item = selected[0]
        values = self.tree.item(self.item)["values"]
        data = dict(zip(self.display_columns, values))
        self.codigo_apertura = data['Codigo_apertura']
        self.codigo_cierre =   data['Codigo_cierre']
        self.fecha_mes = data['Fecha_apertura']

        # Establece el locale en español para que los nombres de los meses salgan en español
        locale.setlocale(locale.LC_TIME, 'Spanish_Spain.1252')  # En Windows

        # Suponiendo que data['Fecha_apertura'] es un string
        fecha_str = data['Fecha_apertura'][:19]  # Recorta los microsegundos extra
        fecha_dt = datetime.strptime(fecha_str, "%Y-%m-%d %H:%M:%S")

        # Extraer el mes en palabras
        mes_en_palabras = fecha_dt.strftime('%B')  # Por ejemplo, 'marzo'

        self.fecha_mes = mes_en_palabras

        # Crear ventana personalizada
        respuesta = {"accion": None}

        def confirmar():
            respuesta["accion"] = "revisado"
            ventana.destroy()

        def solicitar_cambio():
            respuesta["accion"] = "cambio"
            ventana.destroy()

        ventana = Toplevel(self.root)
        ventana.title("Revisión")
        ventana.geometry("350x120")
        ventana.grab_set()  # Bloquea interacción con ventana principal

        Label(ventana, text="¿Desea confirmar o solicitar cambio?").pack(pady=10)

        btn_frame = tk.Frame(ventana)
        btn_frame.pack(pady=5)

        Button(btn_frame, text="Confirmar", width=10, command=confirmar).pack(side="left", padx=10)
        Button(btn_frame, text="Cambio", width=10, command=solicitar_cambio).pack(side="right", padx=10)

        self.root.wait_window(ventana)  # Espera hasta que se cierre la ventana

        if respuesta["accion"] is None:
            return  # El usuario cerró la ventana sin hacer nada

        if respuesta["accion"] == "revisado":
            self.revisado_valor = f"{self.user} - revisado"
            self.revisado = 'revisado'
            self.pendiente = None
            #self.revisado_valor = self.ocupa_evidencia()
            comparativo = self.revisado_valor
            fecha_revision = str(datetime.now().strftime('%Y-%m-%d %H:%M:%S'))
            self.revisado_valor = f"{self.revisado_valor} - {fecha_revision}"
            try:
                if comparativo == f"{self.user} - revisado":
                    cursor = self.conn.cursor()
                    query = f"""
                        UPDATE {self.table_name}
                        SET Revisado_sector = ?, Revisado_operaciones = ?,Revision = ?
                        WHERE Codigo_apertura = ? AND Codigo_cierre = ?
                    """
                    cursor.execute(query, (
                        self.revisado_valor,
                        self.pendiente,
                        self.revisado,
                        str(data['Codigo_apertura']),
                        str(data['Codigo_cierre'])
                    ))
                    self.conn.commit()

                    if self.tree.exists(self.item) and self.active_tab.get() != "Confirmado":
                        self.tree.delete(self.item)
                    ventana.destroy()
                    self.edit_window.destroy()
            except Exception as e:
                self.conn.rollback()
                messagebox.showerror("Error", f"Error al revisar registro: {e}")
                ventana.destroy()
        else:
            self.soporte_comentario()
            return  # Detenemos aquí para esperar la acción del usuario


    def create_login_sharepoint_interface(self):
        if self.sharepoint_user != '':
            self.create_update_data_interface()
        else:
            self.login_shp_window = ctk.CTkToplevel(self.root)
            self.login_shp_window.title("Iniciar sesión - Sharepoint")
            self.login_shp_window.geometry("500x250")
            self.login_shp_window.grab_set()

            center_frame = ctk.CTkFrame(self.login_shp_window, fg_color="transparent")
            center_frame.pack(expand=True)  # Esto asegura que el frame ocupe todo el espacio vertical disponible

            ctk.CTkLabel(center_frame, 
                        text="Ingresar credenciales de Sharepoint", 
                        font=("Arial", 20, "bold")).pack(pady=20)
            
            self.entries_shp = {}
            user_pass_inputs = [
                ("Usuario", "user_entry"),
                ("Contraseña", "pass_entry")
            ]
            
            for text, name in user_pass_inputs:
                defaul_value = tk.StringVar(value="@eneeutcd.hn" if name=="user_entry" else "")
                frame = ctk.CTkFrame(center_frame, fg_color="transparent")
                frame.pack(pady=5)
                
                ctk.CTkLabel(frame, 
                            text=text + ":", 
                            width=140, 
                            anchor="e").pack(side="left", padx=5)
                
                entry = ctk.CTkEntry(frame, width=250, show="*" if name == "pass_entry" else "", textvariable=defaul_value)
                entry.pack(side="left")
                # Bindear Enter key a validate_and_connect
                entry.bind("<Return>", lambda event: self.validate_and_connect_sharepoint())
                
                self.entries_shp[name] = entry
                
                # Botón para mostrar/ocultar contraseña
                if name == "pass_entry":
                    self.show_password = False
                    
                    def toggle_password():
                        self.show_password = not self.show_password
                        entry.configure(show="" if self.show_password else "*")
                        toggle_btn.configure(text="🔓" if self.show_password else "🔒")  # Cambia el ícono
                    
                    toggle_btn = ctk.CTkButton(frame, text="🔒", width=40, command=toggle_password)
                    toggle_btn.pack(side="left", padx=5)
            
            btn_frame = ctk.CTkFrame(center_frame, fg_color="transparent")
            btn_frame.pack(pady=20)
            
            ctk.CTkButton(btn_frame, 
                        text="Ingresar", 
                        command=self.validate_and_connect_sharepoint,
                        fg_color="#2E8B57",
                        hover_color="#245c3d",
                        corner_radius=8,
                        font=("Arial", 12, "bold")).pack(side="left", padx=10)
    
    def validate_and_connect_sharepoint(self):
        user = self.entries_shp['user_entry'].get().strip()
        password = self.entries_shp['pass_entry'].get().strip()
        
        if not user or not password:
            messagebox.showerror("Error de conexión", "Por favor, complete todos los campos.")
            return
        
        self.connect_to_sharepoint()
    
    
    
    def connect_to_sharepoint(self):
        try:
            # Obtener credenciales del formulario
            user = self.entries_shp['user_entry'].get().strip()
            password = self.entries_shp['pass_entry'].get().strip()

            # Crear contexto de autenticación clásico
            ctx_auth = AuthenticationContext(self.sharepoint_url)
            if ctx_auth.acquire_token_for_user(user, password):
                self.ctx = ClientContext(self.sharepoint_url, ctx_auth)
                
                # Probar carga del sitio para confirmar conexión
                site = self.ctx.web
                self.ctx.load(site)
                self.ctx.execute_query()
                
                print("Conexión exitosa a SharePoint con usuario y contraseña.")
                
                # Guardar usuario para referencias internas
                self.sharepoint_user = user
                
                # Cerrar ventana de login si existía
                if hasattr(self, "login_shp_window") and self.login_shp_window:
                    self.login_shp_window.destroy()
                
                # Continuar con la interfaz de actualización de datos
                self.create_update_data_interface()
                
                return self.ctx
            else:
                messagebox.showerror(
                    "Error de autenticación",
                    f"Error en la autenticación: {ctx_auth.get_last_error()}"
                )
                return None

        except Exception as e:
            messagebox.showerror("Error de conexión", f"Error al conectar a SharePoint: {e}")
            return None


    def create_update_data_interface(self):
        print(self.soporte)
        self.update_window = ctk.CTkToplevel(self.root)
        self.update_window.title("Subir documentos a SharePoint")
        self.update_window.geometry("600x500")
        self.update_window.grab_set()

        ctk.CTkLabel(self.update_window, text="Seleccione los archivos a subir:", font=("Arial", 14, "bold")).pack(pady=10)

        file_frame = ctk.CTkFrame(self.update_window, fg_color="transparent")
        file_frame.pack(pady=5)

        self.selected_files = []

        def seleccionar_archivos():
            files = filedialog.askopenfilenames(title="Seleccionar archivos")
            if files:
                self.selected_files = list(files)
                file_listbox.delete(0, tk.END)
                for file in self.selected_files:
                    file_listbox.insert(tk.END, os.path.basename(file))

        if self.soporte == 'No':
            # Comentario
            ctk.CTkLabel(self.update_window, text="Comentario del sector:", font=("Arial", 14, "bold")).pack(pady=10)
            self.comentario_text = ctk.CTkTextbox(self.update_window, width=500, height=80)
            self.comentario_text.pack()
            ctk.CTkButton(self.update_window, text="Subir comentario", command=self.subir_archivos_a_sharepoint).pack(pady=20)
        elif self.soporte == 'Si':
            ctk.CTkButton(file_frame, text="Seleccionar archivos", command=seleccionar_archivos).pack(pady=5)
            file_listbox = tk.Listbox(file_frame, height=6, width=60)
            file_listbox.pack()

            # Comentario
            ctk.CTkLabel(self.update_window, text="Comentario del sector:", font=("Arial", 14, "bold")).pack(pady=10)
            self.comentario_text = ctk.CTkTextbox(self.update_window, width=500, height=80)
            self.comentario_text.pack()
            ctk.CTkButton(self.update_window, text="Subir archivos y comentario", command=self.subir_archivos_a_sharepoint).pack(pady=20)
            

    # def create_update_data_interface(self):
    #     self.update_window = ctk.CTkToplevel(self.root)
    #     self.update_window.title("Subir documentos a SharePoint")
    #     self.update_window.geometry("600x500")
    #     self.update_window.grab_set()

    #     ctk.CTkLabel(self.update_window, text="Seleccione los archivos a subir:", font=("Arial", 14, "bold")).pack(pady=10)

    #     file_frame = ctk.CTkFrame(self.update_window, fg_color="transparent")
    #     file_frame.pack(pady=5)

    #     self.selected_files = []

    #     def seleccionar_archivos():
    #         files = filedialog.askopenfilenames(title="Seleccionar archivos")
    #         if files:
    #             self.selected_files = list(files)
    #             file_listbox.delete(0, "end")
    #             for file in self.selected_files:
    #                 file_listbox.insert("end", os.path.basename(file))

    #     ctk.CTkButton(file_frame, text="Seleccionar archivos", command=seleccionar_archivos).pack(pady=5)
    #     file_listbox = tk.Listbox(file_frame, height=6, width=60)
    #     file_listbox.pack()

    #     # Comentario
    #     ctk.CTkLabel(self.update_window, text="Comentario del sector:", font=("Arial", 14, "bold")).pack(pady=10)
    #     self.comentario_text = ctk.CTkTextbox(self.update_window, width=500, height=80)
    #     self.comentario_text.pack()
    #     ctk.CTkButton(self.update_window, text="Subir archivos y comentario", command=self.subir_archivos_a_sharepoint).pack(pady=20)


    

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

    def seleccionar_archivos(self):
        archivos = filedialog.askopenfilenames(
            title="Seleccionar archivos",
            filetypes=[("Todos los archivos", "*.*")]
        )
        self.selected_files = list(archivos)
        if self.selected_files:
            messagebox.showinfo("Archivos seleccionados", f"{len(self.selected_files)} archivo(s) seleccionado(s).")
        else:
            messagebox.showwarning("Advertencia", "No se seleccionaron archivos.")

    # def subir_archivos_a_sharepoint(self):
    #     if self.soporte == 'No':
    #         comentario = self.comentario_text.get("1.0", "end").strip()
    #         self.comen = comentario  # Guardas el comentario limpio
    #         self.update_window.destroy()
    #         fecha_revision = str(datetime.now().strftime('%Y-%m-%d %H:%M:%S'))
    #         #self.revisado_valor = f"{self.user} - solicitó cambio - {fecha_revision}"
            
    #         try:
    #             cursor = self.conn.cursor()
    #             query = f"""
    #                 UPDATE {self.table_name}
    #                 SET Revisado_sector = ?, Comentario_sector = ?, Revisado_operaciones = ?, Revision = ?
    #                 WHERE Codigo_apertura = ? AND Codigo_cierre = ?
    #             """
    #             cursor.execute(query, (
    #                 self.revisado_valor,
    #                 self.comen,
    #                 self.pendiente,
    #                 self.revisado,
    #                 str(self.codigo_apertura),
    #                 str(self.codigo_cierre)
    #             ))
    #             self.conn.commit()
    #             #return
    #             if self.tree.exists(self.item) and self.active_tab.get() != "Confirmado":
    #                 self.tree.delete(self.item)
    #             messagebox.showinfo("Éxito", "Cambios guardados correctamente")
    #             return
    #         except Exception as e:
    #             print(f"Error al subir archivos: {e}")
    #             messagebox.showerror("Error", f"No se pudieron subir los archivos: {e}")
    #             self.update_window.destroy()


    #     if not self.selected_files:
    #         messagebox.showwarning("Advertencia", "No hay archivos seleccionados.")
    #         return

    #     try:
    #         cod_apertura = str(self.codigo_apertura)
    #         cod_cierre = str(self.codigo_cierre)
    #         carpeta_nombre = f"{cod_apertura}_{cod_cierre}"
    #         carpeta_mes = str(self.fecha_mes)
    #         print(f'carpeta mes es : {carpeta_mes}')

    #         base_path = "/sites/ControlGestin/Documentos compartidos/PRUEBAS"
    #         carpeta_intermedia = self.carpetas  # Ej: "TEGUCIGALPA"
                
    #         subfolders = [carpeta_intermedia,carpeta_mes,carpeta_nombre]

    #         # Verifica o crea la ruta completa
    #         target_folder,_ = self.ensure_folder_path_sharepoint(self.ctx, base_path, subfolders)
    #         print(target_folder)
    #         _, self.current_path = self.ensure_folder_path_sharepoint(self.ctx, base_path, subfolders)
            
    #         # Subir todos los archivos seleccionados
    #         for archivo in self.selected_files:
    #             with open(archivo, "rb") as f:
    #                 contenido = f.read()
    #             nombre_archivo = os.path.basename(archivo)
    #             target_folder.upload_file(nombre_archivo, contenido).execute_query()

    #         messagebox.showinfo("Éxito", f"{len(self.selected_files)} archivo(s) subido(s) correctamente.")
            

    #         comentario = self.comentario_text.get("1.0", "end").strip()
    #         self.comen = comentario  # Guardas el comentario limpio
    #         self.update_window.destroy()
    #         fecha_revision = str(datetime.now().strftime('%Y-%m-%d %H:%M:%S'))
    #         #self.revisado_valor = f"{self.user} - solicitó cambio - {fecha_revision}"
            
    #         try:
    #             cursor = self.conn.cursor()
    #             query = f"""
    #                 UPDATE {self.table_name}
    #                 SET Revisado_sector = ?, Comentario_sector = ?, Revisado_operaciones = ?, Revision = ?
    #                 WHERE Codigo_apertura = ? AND Codigo_cierre = ?
    #             """
    #             cursor.execute(query, (
    #                 self.revisado_valor,
    #                 self.comen,
    #                 self.pendiente,
    #                 self.revisado,
    #                 str(self.codigo_apertura),
    #                 str(self.codigo_cierre)
    #             ))
    #             self.conn.commit()



    #             self.update_window.destroy()
    #             if self.tree.exists(self.item) and self.active_tab.get() != "Confirmado":
    #                 self.tree.delete(self.item)
    #                # messagebox.showinfo("Éxito", "Cambios guardados correctamente")
    #         except Exception as e:
    #             self.conn.rollback()
    #             messagebox.showerror("Error", f"Error al guardar cambios: {e}")
            
    #     except Exception as e:
    #         print(f"Error al subir archivos: {e}")
    #         messagebox.showerror("Error", f"No se pudieron subir los archivos: {e}")


    # def subir_archivos_a_sharepoint(self):
    #     FLOW_CREATE_FOLDER = "https://defaultc1b22713e01544978af7ac76803fda.c5.environment.api.powerplatform.com:443/powerautomate/automations/direct/workflows/0c90a9e0c1f14a05968eb7ec0458b029/triggers/manual/paths/invoke/?api-version=1&sp=%2Ftriggers%2Fmanual%2Frun&sv=1.0&sig=_k1Nun1ef72pYVccwhCO8j-T85WOV0aLqFHHEgKSIjc"

    #     FLOW_UPLOAD_FILE = "https://defaultc1b22713e01544978af7ac76803fda.c5.environment.api.powerplatform.com:443/powerautomate/automations/direct/workflows/9176607fb1514856b5a5e3855c54ce7c/triggers/manual/paths/invoke/?api-version=1&sp=%2Ftriggers%2Fmanual%2Frun&sv=1.0&sig=8Cu_HQ-eGOREtqZosF3iljZ-FqeLjmtQgNN61hncvmQ"
    #     if not self.selected_files:
    #         messagebox.showwarning("Advertencia", "No hay archivos seleccionados.")
    #         return

    #     carpeta_nombre = f"{self.codigo_apertura}_{self.codigo_cierre}"
    #     carpeta_mes = str(self.fecha_mes)
    #     sector = self.carpetas

    #     # 1️⃣ Crear carpeta usando flujo
    #     payload_folder = {
    #         "sector": sector,
    #         "mes": carpeta_mes,
    #         "carpeta_final": carpeta_nombre
    #     }
    #     headers = {"Content-Type": "application/json"}
    #     response_folder = requests.post(FLOW_CREATE_FOLDER, json=payload_folder, headers=headers)

    #     if response_folder.status_code != 200:
    #         messagebox.showerror("Error", f"No se pudo crear la carpeta: {response_folder.text}")
    #         return

    #     # 2️⃣ Subir archivos
    #     for archivo in self.selected_files:
    #         with open(archivo, "rb") as f:
    #             file_bytes = f.read()
    #             file_base64 = base64.b64encode(file_bytes).decode("utf-8")

    #         payload_file = {
    #             "folderPath": f"Documentos compartidos/PRUEBAS/{sector}/{carpeta_mes}/{carpeta_nombre}",
    #             "filename": os.path.basename(archivo),
    #             "filecontent": file_base64
    #         }
    #         print(f"aqui esta la ruta Documentos compartidos/PRUEBAS/{sector}/{carpeta_mes}/{carpeta_nombre}")
    #         response_file = requests.post(FLOW_UPLOAD_FILE, json=payload_file, headers=headers)

    #         if response_file.status_code != 200:
    #             messagebox.showerror("Error", f"No se pudo subir {archivo}: {response_file.text}")
    #             return

    #     messagebox.showinfo("Éxito", f"{len(self.selected_files)} archivo(s) subido(s) correctamente.")

    #     # Guardar comentario en base de datos
    #     comentario = self.comentario_text.get("1.0", "end").strip()
    #     self.comen = comentario
    #     self.update_window.destroy()

    #     try:
    #         cursor = self.conn.cursor()
    #         query = f"""
    #             UPDATE {self.table_name}
    #             SET Revisado_sector = ?, Comentario_sector = ?, Revisado_operaciones = ?, Revision = ?
    #             WHERE Codigo_apertura = ? AND Codigo_cierre = ?
    #         """
    #         cursor.execute(query, (
    #             self.revisado_valor,
    #             self.comen,
    #             self.pendiente,
    #             self.revisado,
    #             str(self.codigo_apertura),
    #             str(self.codigo_cierre)
    #         ))
    #         self.conn.commit()
    #         if self.tree.exists(self.item):
    #             self.tree.delete(self.item)
    #     except Exception as e:
    #         self.conn.rollback()
    #         messagebox.showerror("Error", f"Error al guardar cambios: {e}"


    def subir_archivos_a_sharepoint(self):
        FLOW_CREATE_FOLDER = "https://defaultc1b22713e01544978af7ac76803fda.c5.environment.api.powerplatform.com:443/powerautomate/automations/direct/workflows/0c90a9e0c1f14a05968eb7ec0458b029/triggers/manual/paths/invoke/?api-version=1&sp=%2Ftriggers%2Fmanual%2Frun&sv=1.0&sig=_k1Nun1ef72pYVccwhCO8j-T85WOV0aLqFHHEgKSIjc"
        FLOW_UPLOAD_FILE = "https://defaultc1b22713e01544978af7ac76803fda.c5.environment.api.powerplatform.com:443/powerautomate/automations/direct/workflows/9176607fb1514856b5a5e3855c54ce7c/triggers/manual/paths/invoke/?api-version=1&sp=%2Ftriggers%2Fmanual%2Frun&sv=1.0&sig=8Cu_HQ-eGOREtqZosF3iljZ-FqeLjmtQgNN61hncvmQ"

        # =========================
        # Caso soporte "No" (solo comentario)
        # =========================
        if self.soporte == 'No':
            comentario = self.comentario_text.get("1.0", "end").strip()
            self.comen = comentario
            self.update_window.destroy()

            try:
                cursor = self.conn.cursor()
                query = f"""
                    UPDATE {self.table_name}
                    SET Revisado_sector = ?, Comentario_sector = ?, Revisado_operaciones = ?, Revision = ?
                    WHERE Codigo_apertura = ? AND Codigo_cierre = ?
                """
                cursor.execute(query, (
                    self.revisado_valor,
                    self.comen,
                    self.pendiente,
                    self.revisado,
                    str(self.codigo_apertura),
                    str(self.codigo_cierre)
                ))
                self.conn.commit()

                if self.tree.exists(self.item) and self.active_tab.get() != "Confirmado":
                    self.tree.delete(self.item)
                messagebox.showinfo("Éxito", "Cambios guardados correctamente")
                return
            except Exception as e:
                print(f"Error al guardar comentario: {e}")
                messagebox.showerror("Error", f"No se pudieron guardar los cambios: {e}")
                self.update_window.destroy()
                return

        # =========================
        # Validación archivos seleccionados
        # =========================
        if not self.selected_files:
            messagebox.showwarning("Advertencia", "No hay archivos seleccionados.")
            return

        # =========================
        # Crear carpeta usando flujo
        # =========================
        carpeta_nombre = f"{self.codigo_apertura}_{self.codigo_cierre}"
        carpeta_mes = str(self.fecha_mes)
        sector = self.carpetas

        payload_folder = {
            "sector": sector,
            "mes": carpeta_mes,
            "carpeta_final": carpeta_nombre
        }
        headers = {"Content-Type": "application/json"}

        response_folder = requests.post(FLOW_CREATE_FOLDER, json=payload_folder, headers=headers)
        if response_folder.status_code != 200:
            messagebox.showerror("Error", f"No se pudo crear la carpeta: {response_folder.text}")
            return

        # =========================
        # Subir archivos uno por uno
        # =========================
        errores = []
        for archivo in self.selected_files:
            try:
                with open(archivo, "rb") as f:
                    file_bytes = f.read()
                    file_base64 = base64.b64encode(file_bytes).decode("utf-8")

                payload_file = {
                    "folderPath": f"Documentos compartidos/PRUEBAS/{sector}/{carpeta_mes}/{carpeta_nombre}",
                    "filename": os.path.basename(archivo),
                    "filecontent": file_base64
                }

                response_file = requests.post(FLOW_UPLOAD_FILE, json=payload_file, headers=headers)
                if response_file.status_code != 200:
                    errores.append(f"{archivo}: {response_file.text}")

            except Exception as e:
                errores.append(f"{archivo}: {str(e)}")

        if errores:
            messagebox.showerror("Error", f"Algunos archivos no se pudieron subir:\n" + "\n".join(errores))
            return

        # =========================
        # Guardar comentario y actualizar BD
        # =========================
        comentario = self.comentario_text.get("1.0", "end").strip()
        self.comen = comentario
        self.update_window.destroy()

        try:
            cursor = self.conn.cursor()
            query = f"""
                UPDATE {self.table_name}
                SET Revisado_sector = ?, Comentario_sector = ?, Revisado_operaciones = ?, Revision = ?
                WHERE Codigo_apertura = ? AND Codigo_cierre = ?
            """
            cursor.execute(query, (
                self.revisado_valor,
                self.comen,
                self.pendiente,
                self.revisado,
                str(self.codigo_apertura),
                str(self.codigo_cierre)
            ))
            self.conn.commit()

            if self.tree.exists(self.item) and self.active_tab.get() != "Confirmado":
                self.tree.delete(self.item)

            messagebox.showinfo("Éxito", f"{len(self.selected_files)} archivo(s) subido(s) correctamente.")

        except Exception as e:
            self.conn.rollback()
            messagebox.showerror("Error", f"Error al guardar cambios: {e}")


    def abrir_si_existe_sharepoint(self,cod_apertura, cod_cierre, carpeta_intermedia,carpeta_mes):
        

        carpeta_nombre = f"{cod_apertura}_{cod_cierre}"
        base_path = "/sites/ControlGestin/Documentos compartidos/PRUEBAS"
        current_path = f"{base_path}/{carpeta_intermedia}/{carpeta_mes}/{carpeta_nombre}"
        print(f'fecha = {current_path}')

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


    def ensure_folder_path_sharepoint(self, ctx, root_path, subfolders):
        """
        Asegura que todas las carpetas en la ruta de SharePoint existen. Si no existen, las crea.

        :param ctx: El contexto de SharePoint.
        :param root_path: Ruta base en SharePoint (ej. /sites/ControlGestin/Documentos compartidos/PRUEBAS).
        :param subfolders: Lista de subcarpetas a crear dentro del root_path.
        :return: Objeto de la última carpeta en la ruta completa.
        """
        current_path = root_path

        for subfolder in subfolders:
            current_path = f"{current_path}/{subfolder}"  # Concatenación en formato SharePoint (no usar os.path)
            folder = ctx.web.get_folder_by_server_relative_url(current_path)
            try:
                ctx.load(folder)
                ctx.execute_query()
                print(f"La carpeta ya existe: {current_path}")
            except:
                # Crear carpeta si no existe
                parent_folder = ctx.web.get_folder_by_server_relative_url(os.path.dirname(current_path))
                parent_folder.folders.add(subfolder).execute_query()
                print(f"Carpeta creada: {current_path}")
                folder = ctx.web.get_folder_by_server_relative_url(current_path)
                ctx.load(folder)
                ctx.execute_query()

        return folder,current_path


    def edit_row(self):
        selected = self.tree.selection()
        if not selected:
            messagebox.showwarning("Advertencia", "Seleccione un registro para editar")
            return
        self.soporte == 'Si'
        self.share = 'No'
        self.item = selected[0]
        values = self.tree.item(self.item)["values"]
        data = dict(zip(self.display_columns, values))

        self.codigo_apertura = data['Codigo_apertura']
        self.codigo_cierre =   data['Codigo_cierre']
        self.revi = data['Revisado_sector']
        print(f'revisado sectttttttttttttttttor es {self.revi}')

        # Establece el locale en español para que los nombres de los meses salgan en español
        locale.setlocale(locale.LC_TIME, 'Spanish_Spain.1252')  # En Windows

        # Suponiendo que data['Fecha_apertura'] es un string
        fecha_str = data['Fecha_apertura'][:19]  # Recorta los microsegundos extra
        fecha_dt = datetime.strptime(fecha_str, "%Y-%m-%d %H:%M:%S")

        # Extraer el mes en palabras
        mes_en_palabras = fecha_dt.strftime('%B')  # Por ejemplo, 'marzo'

        self.fecha_mes = mes_en_palabras
   
        self.sectores_ = ['DANLI','TEGUCIGALPA','CHOLUTECA','COMAYAGUA',
                    'JUTICALPA','TOCOA','LA CEIBA','SAN PEDRO SULA',
                    'VILLANUEVA','SANTA ROSA DE COPAN',
                    'SANTA CRUZ DE YOJOA','EL PROGRESO']

        aux = data['Sector']
        self.carpetas = aux
        print(self.carpetas)
        print(self.todos)
        
        if 'con soporte' in data['Revisado_sector'].lower() or 'adjunta soporte' in data['Revisado_sector'].lower():
            print('aqui')
            self.share = 'Si'
        self.edit_window = ctk.CTkToplevel(self.root)
        self.edit_window.title("Editar Registro")
        self.edit_window.geometry("620x720")
        self.edit_window.grab_set()

        entries = {}

        # Crear scrollable frame
        scrollable_frame = ctk.CTkScrollableFrame(self.edit_window, width=600, height=650)
        scrollable_frame.pack(padx=10, pady=10, fill="both", expand=True)

        read_only_fields = [
            'Codigo_apertura', 'Codigo_cierre', 'Fecha_apertura', 'Fecha_cierre',
            'Tiempo_horas', 'Tiempo_minutos', 'Zona', 'Sector', 'Subestacion',
            'Circuito', 'Grupo_calidad', 'Tipo_interruptor', 'Equipo_opero', 'Ubicacion',
            'Carga_MVA', 'Relevador', 'Interrupcion', 'Clasificacion',
            'Registro_interrupcion', 'Estado', 'Revisado_sector','Revisado_operaciones'
        ]

        for col in self.editable_columns:
            if col in ('Comentario_operaciones','Comentario_sector','Revisado_operaciones')and self.revi == '-':
                continue 
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
                continue

            if col in ["Observacion", "Comentario_operaciones",'Comentario_sector']:
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
            text="Revisión sector",
            command=self.revisar_row,
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
        
        if self.active_tab.get() != 'pendiente' and  self.share == 'Si':
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

        #self.edit_window.bind("<Return>", lambda e: self.save_changes(self.item, entries, self.edit_window))


    def save_changes(self, item, entries, window):
        original_values = self.tree.item(self.item)["values"]
        original_data = dict(zip(self.display_columns, original_values))
        
        # Campos de solo lectura
        read_only_fields = ['Circuito', 'Subestacion', 'Ubicacion']
        
        # Construir new_data excluyendo campos bloqueados
        new_data = {}
        for col in self.editable_columns:
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
                
            elif col == "Observacion":
                new_data[col] = entries[col].get("1.0", "end-1c").strip()
                
            else:
                new_data[col] = entries[col].get().strip()
        
        new_data['Usuario_actualizacion'] = self.user

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

        if new_data['Tiempo_horas'] <= 0.05:
            new_data['Registro_interrupcion'] = 17
            new_data['Interrupcion'] = 'INSTANTANEA'
            new_data['Clasificacion'] = 'E'
        # Validación de tipos de datos
        type_errors = []
        for col in self.editable_columns:
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
            for col in self.editable_columns if col not in read_only_fields
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
                            Saifi_contribucion_global = CAST(Clientes_afectados AS FLOAT) / NULLIF(CAST(Clientes_nacional AS FLOAT), 0),
                            Saidi_contribucion_global = (CAST(Clientes_afectados AS FLOAT) * CAST(Tiempo_horas AS FLOAT)) / NULLIF(CAST(Clientes_nacional AS FLOAT), 0),
                            Saifi_grupo = CAST(Clientes_afectados AS FLOAT) / NULLIF(CAST(Clientes_grupo AS FLOAT), 0),
                            Saidi_grupo = (CAST(Clientes_afectados AS FLOAT) * CAST(Tiempo_horas AS FLOAT)) / NULLIF(CAST(Clientes_grupo AS FLOAT), 0),
                            Saifi_zona = CAST(Clientes_afectados AS FLOAT) / NULLIF(CAST(Clientes_zona AS FLOAT), 0),
                            Saidi_zona = (CAST(Clientes_afectados AS FLOAT) * CAST(Tiempo_horas AS FLOAT)) / NULLIF(CAST(Clientes_zona AS FLOAT), 0),
                            Saifi_sector = CAST(Clientes_afectados AS FLOAT) / NULLIF(CAST(Clientes_sector AS FLOAT), 0),
                            Saidi_sector = (CAST(Clientes_afectados AS FLOAT) * CAST(Tiempo_horas AS FLOAT)) / NULLIF(CAST(Clientes_sector AS FLOAT), 0)
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
        self.item = selected[0]
        values = self.tree.item(self.item, "values")
        
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