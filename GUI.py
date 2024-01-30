import tkinter as tk   # Carga módulo tk (widgets estándar)
from tkinter import messagebox
from tkinter import ttk  # Carga ttk (para widgets nuevos 8.5+)
from Core.BancoDeMediciones import *
from Core.ATV import *
from Core.DTV import *


# Define la ventana principal de la aplicación
class WelcomeWindow():
    def __init__(self) -> None:

        # Creación de la ventana
        self.window = tk.Tk()
        self.window.iconphoto(False, tk.PhotoImage(file="./Imagenes/logo-ane-16.png"))
        self.window.title('Proyecto Tecno FSH-ETL')
        
        # Dimensiones de la ventana
        window_width = 426
        window_height = 240
        
        # Obtener las dimensiones de la pantalla
        screen_width = self.window.winfo_screenwidth()
        screen_height = self.window.winfo_screenheight()
        
        # Calcular la posición para centrar la ventana en la pantalla
        x = (screen_width // 2) - (window_width // 2)
        y = (screen_height // 2) - (window_height // 2)
        
        # Configurar el redimensionamiento de la ventana
        self.window.columnconfigure(0, weight=1)
        self.window.rowconfigure(0, weight=1)
        
        # Configurar la geometría de la ventana
        self.window.geometry(f"{window_width}x{window_height}+{x}+{y}")

        # Creación del Frame
        self.button_frame = ttk.Frame(self.window)
        self.button_frame.pack(expand=True)

        # Creación de los botones
        button_atv = ttk.Button(self.button_frame, text='Medición de TV analógica', command=self.open_atv)
        button_dtv = ttk.Button(self.button_frame, text='Medición de TV digital', command=self.open_dtv)
        button_bm  = ttk.Button(self.button_frame, text='Banco de mediciones', command=self.open_bm)

        # Posicionamiento de los botones
        button_atv.grid(row=0, column=0, padx=10, pady=20)
        button_dtv.grid(row=1, column=0, padx=10, pady=20)
        button_bm.grid(row=2, column=0, padx=10, pady=20)

        self.window.mainloop()

    # Abrir la ventana de televisión analógica
    def open_atv(self):
        self.window.destroy()
        ATVWindow()
    
    #Abrir la ventana de televisión digital
    def open_dtv(self):
        self.window.destroy()
        DTVWindow()

    # Abrir la ventana de banco de mediciones
    def open_bm(self):
        self.window.destroy()
        BancoMedicionesWindow()


# Define la ventana 'Banco de mediciones'
class BancoMedicionesWindow():
    def __init__(self) -> None:
        # Creación de la ventana
        self.window = tk.Tk()
        self.window.iconphoto(False, tk.PhotoImage(file="./Imagenes/logo-ane-16.png"))
        self.window.title('Banco de mediciones')
        
        # Dimensiones de la ventana
        window_width = 426
        window_height = 240
        
        # Obtener las dimensiones de la pantalla
        screen_width  = self.window.winfo_screenwidth()
        screen_height = self.window.winfo_screenheight()
        
        # Calcular la posición para centrar la ventana en la pantalla
        x = (screen_width // 2) - (window_width // 2)
        y = (screen_height // 2) - (window_height // 2)
        
        # Configurar el redimensionamiento de la ventana
        self.window.columnconfigure(0, weight=1)
        self.window.rowconfigure(0, weight=1)
        
        # Configurar la geometría de la ventana
        self.window.geometry(f"{window_width}x{window_height}+{x}+{y}")

        # Creación de los Frames
        self.conection_frame   = ttk.Frame(self.window)
        self.ckecklist_frame   = ttk.Frame(self.window)
        self.measurement_frame = ttk.Frame(self.window)

        self.conection_frame.pack(fill=tk.X)
        self.ckecklist_frame.pack(fill=tk.X)
        self.measurement_frame.pack(fill=tk.X)

        # Creación del Entry para la IP del instrumento
        self.ip_adress_entry = ttk.Entry(self.conection_frame)
        self.ip_adress_entry.grid(row=0, column=0, padx=10, pady=10)

        # Creación del botón para conectar el dispositivo
        self.button_probe = ttk.Button(self.conection_frame, text='Probar conexión', command=self.probe_conection)
        self.button_probe.grid(row=0, column=1, padx=10, pady=10)

        # Creación del label de estado de la conexión
        self.conection_text = tk.StringVar()
        self.conection_text.set("Esperando conexión")
        self.label_status = ttk.Label(self.conection_frame, textvariable=self.conection_text)
        self.label_status.grid(row=1, column=0, padx=10, pady=5)

        # Nombre de las bandas que disponibles para la medida
        name_bands = ['700','850','1900','AWS1','AWS2','AWS3','2.5','900_1','900_2','3500','2.4GHz','5GHz','2300MHz','FM']

        # Creación del diccionario necesario para crear los check buttons
        self.bands = dict()
        for name in name_bands:
            self.bands[name] = tk.BooleanVar()

        # Creación de los check buttons
        i = 0
        for key, var in self.bands.items():
            check_button = ttk.Checkbutton(self.ckecklist_frame, text=key, variable=var)
            check_button.grid(row=i//4, column=i%4, padx=10, pady=2, sticky='w')
            i += 1
    
        # Creación del botón para iniciar la medida    
        self.button_start = ttk.Button(self.measurement_frame, text='Iniciar medición', command=self.start_measurement)
        self.button_start.grid(row=0, column=0, padx=10, pady=10)

        # Creación del label de estado de la conexión
        self.measure_text = tk.StringVar()
        self.measure_text.set('Esperando medición')
        self.label_measure = ttk.Label(self.measurement_frame, textvariable=self.measure_text)
        self.label_measure.grid(row=0, column=1, padx=10, pady=10)

        # Obtener la altura requerida para los nuevos widgets
        self.window.update_idletasks()
        new_width  = self.window.winfo_reqwidth()
        new_height = self.window.winfo_reqheight()

        self.window.geometry(f"{new_width}x{new_height}")

        self.window.mainloop()

    # Función del botón 'Probar conexión'
    def probe_conection(self):
        # Obtención de la IP ingresada por el usuario
        # En caso de estar vacía, la rutina buscará los dispositivos disponibles
        ip_adress = self.ip_adress_entry.get()

        # Conexión con el dispositivo
        try:
            self.fsh = BancoDeMediciones(ip_adress, True, True, "Simulate=True")
            self.fsh.instrument_status_checking = True  # Error check after each command
            self.fsh.data_chunk_size = 100 # Definición del tamaño del buffer
            idn = self.fsh.idn_string.split(sep=',')[1]
            self.conection_text.set(f'Conectado a: FSH4')
        
        # En caso de no lograrse, se le notifica al usuario en el Label
        except:
            self.conection_text.set('Error en la conexión')

        # Comandos iniciales para el instrumento
        self.fsh.setup()

    # Función del botón 'Iniciar medición'
    def start_measurement(self):

        # Creación de carpeta de soportes
        self.fsh.create_folder()

        # Bandas que se medirán
        bands = []

        # Se crea el listado de bandas a medir, según los check button
        for text, var in self.bands.items():
            if var.get() == True:
                bands.append(text)
        
        # Mensaje de aviso para conectar al instrumento la antena utilizada
        # para medir todas las bandas, excepto FM
        if len(bands) > 1:
            messagebox.showinfo(message="Conecte la antena de las otras bandas \n y de clic en aceptar", title='Aviso')

        # Medición de las bandas seleccionadas
        for band in bands:

            # Mensaje de aviso para conectar al instrumento la antena utilizada para FM
            if band == 'FM':
                messagebox.showinfo(message="Conecte la antena de FM \n y de clic en aceptar", title='Aviso')
            
            # Puesta en marcha de la medida de la banda
            self.fsh.measurement(band)

            # Mensaje de que banda se está leyendo en el Label
            self.measure_text.set(f'Midiendo {band}')

        # Mensaje de medición finalizada en el Label
        self.measure_text.set('Medición finalizada')

        # Cierra conexión con el instrumento
        self.fsh.close()


# Define la ventana 'Medición de televisión analógica'
class ATVWindow():
    def __init__(self) -> None:

        self.window = tk.Tk()
        self.window.iconphoto(False, tk.PhotoImage(file="./Imagenes/logo-ane-16.png"))
        self.window.title('Medición de televisión analógica')
        
        # Dimensiones de la ventana
        window_width = 426
        window_height = 240
        
        # Obtener las dimensiones de la pantalla
        screen_width  = self.window.winfo_screenwidth()
        screen_height = self.window.winfo_screenheight()
        
        # Calcular la posición para centrar la ventana en la pantalla
        x = (screen_width // 2) - (window_width // 2)
        y = (screen_height // 2) - (window_height // 2)
        
        # Configurar el redimensionamiento de la ventana
        self.window.columnconfigure(0, weight=1)
        self.window.rowconfigure(0, weight=1)
        
        # Configurar la geometría de la ventana
        self.window.geometry(f"{window_width}x{window_height}+{x}+{y}")

        # Creación de los Frames
        self.conection_frame   = ttk.Frame(self.window)
        self.channels_frame    = ttk.Frame(self.window)
        self.measurement_frame = ttk.Frame(self.window)

        self.conection_frame.pack(fill=tk.X)
        self.channels_frame.pack(fill=tk.X)
        self.measurement_frame.pack(fill=tk.X)


        ## Primer Frame
        # Creación del Entry para la IP del instrumento
        self.ip_adress_entry = ttk.Entry(self.conection_frame)
        self.ip_adress_entry.grid(row=0, column=0, padx=10, pady=10)

        # Creación del botón para conectar el dispositivo
        self.button_probe = ttk.Button(self.conection_frame, text='Probar conexión', command=self.probe_conection)
        self.button_probe.grid(row=0, column=1, padx=10, pady=10)

        # Creación del label de estado de la conexión
        self.conection_text = tk.StringVar()
        self.conection_text.set("Esperando conexión")
        self.label_status = ttk.Label(self.conection_frame, textvariable=self.conection_text)
        self.label_status.grid(row=1, column=0, padx=10, pady=5)


        ## Segundo Frame
        # Creación del label de instrucción para ingreso de número de canales
        self.num_of_channels_label = ttk.Label(self.channels_frame, text='Ingrese el número de canales que se medirán')
        self.num_of_channels_label.grid(row=0, column=0, padx=5, pady=2)

        # Creación del Entry para el ingreso de número de canales
        self.num_of_channels_entry = ttk.Entry(self.channels_frame, width=5)
        self.num_of_channels_entry.grid(row=0, column=1, padx=5, pady=2)

        # Creación del botón para conectar el dispositivo
        self.add_channels_button = ttk.Button(self.channels_frame, text='Crear', command=self.create_list)
        self.add_channels_button.grid(row=0, column=2, padx=5, pady=2)


        ## Tercer Frame
        # Creación del botón para iniciar la medida    
        self.button_start = ttk.Button(self.measurement_frame, text='Iniciar medición', command=self.start_measurement)
        self.button_start.grid(row=0, column=0, padx=10, pady=10)

        # Creación del label de estado de la conexión
        self.measure_text = tk.StringVar()
        self.measure_text.set('Esperando medición')
        self.label_measure = ttk.Label(self.measurement_frame, textvariable=self.measure_text)
        self.label_measure.grid(row=0, column=1, padx=10, pady=10)

        # Obtener la altura requerida para los nuevos widgets
        self.window.update_idletasks()
        new_width  = self.window.winfo_reqwidth()
        new_height = self.window.winfo_reqheight()

        self.window.geometry(f"{new_width}x{new_height}")

        self.window.mainloop()

    # Función del botón 'Probar conexión'
    def probe_conection(self):
        # Obtención de la IP ingresada por el usuario
        # En caso de estar vacía, la rutina buscará los dispositivos disponibles
        ip_adress = self.ip_adress_entry.get()

        # Conexión con el dispositivo
        try:
            self.fsh = ATV(ip_adress, True, True, "Simulate=True")
            self.fsh.instrument_status_checking = True  # Error check after each command
            self.fsh.data_chunk_size = 100 # Definición del tamaño del buffer
            idn = self.fsh.idn_string.split(sep=',')[1]
            self.conection_text.set(f'Conectado a: FSH4')
        
        # En caso de no lograrse, se le notifica al usuario en el Label
        except:
            self.conection_text.set('Error en la conexión')

        # Comandos iniciales para el instrumento
        self.fsh.setup()

    # Función del botón 'Crear'
    def create_list(self):
        number_of_channels = self.num_of_channels_entry.get()
        try:
            number_of_channels = int(number_of_channels)
        except ValueError:
            messagebox.showerror(message='El valor ingresado no es un número')
        
        self.channels = []
        for i in range(number_of_channels):
            channel_label = ttk.Label(self.channels_frame, text=f'Ingrese el canal # {i+1}')
            channel_label.grid(row=i+1, column=0, padx=5, pady=2, sticky='w')
            channel_entry = ttk.Entry(self.channels_frame, width=5)
            channel_entry.grid(row=i+1, column=1, padx=5, pady=2)
            self.channels.append(channel_entry)

        # Obtener la altura requerida para los nuevos widgets
        self.window.update_idletasks()
        new_width  = self.window.winfo_reqwidth()
        new_height = self.window.winfo_reqheight()

        self.window.geometry(f"{new_width}x{new_height}")

    # Función del botón 'Iniciar medición'
    def start_measurement(self):
        
        # Creación de carpeta de soportes
        self.fsh.create_folder()

        for channel in self.channels:
            channel = channel.get()

            try:
                channel = int(channel)
            except ValueError:
                messagebox.showerror(message=f'El valor {channel} no es un número')

            # Mensaje de aviso para apuntar la antena hacia el acimuth deseado
            messagebox.showinfo(message=f'Apunte la antena hacia el acimuth del canal {channel} y luego de clic en aceptar')

            # Puesta en marcha de la medida de la banda
            self.fsh.measurement(channel)

            # Mensaje de que banda se está leyendo en el Label
            self.measure_text.set(f'Midiendo el canal {channel}')

        # Mensaje de medición finalizada en el Label
        self.measure_text.set('Medición finalizada')

        # Cierra conexión con el instrumento
        self.fsh.close()


# Define la ventana 'Medición de televisión digital'
class DTVWindow():
    def __init__(self) -> None:
        self.window = tk.Tk()
        self.window.iconphoto(False, tk.PhotoImage(file="./Imagenes/logo-ane-16.png"))
        self.window.title('Medición de televisión digital')
        
        # Dimensiones de la ventana
        window_width = 426
        window_height = 240
        
        # Obtener las dimensiones de la pantalla
        screen_width  = self.window.winfo_screenwidth()
        screen_height = self.window.winfo_screenheight()
        
        # Calcular la posición para centrar la ventana en la pantalla
        x = (screen_width // 2) - (window_width // 2)
        y = (screen_height // 2) - (window_height // 2)
        
        # Configurar el redimensionamiento de la ventana
        self.window.columnconfigure(0, weight=1)
        self.window.rowconfigure(0, weight=1)
        
        # Configurar la geometría de la ventana
        self.window.geometry(f"{window_width}x{window_height}+{x}+{y}")

        # Creación de los Frames
        self.conection_frame   = ttk.Frame(self.window)
        self.parameters_frame  = ttk.Frame(self.window)
        self.channels_frame    = ttk.Frame(self.window)
        self.measurement_frame = ttk.Frame(self.window)

        self.conection_frame.pack(fill=tk.X)
        self.parameters_frame.pack(fill=tk.X)
        self.channels_frame.pack(fill=tk.X)
        self.measurement_frame.pack(fill=tk.X)


        ## Primer Frame
        # Creación del Entry para la IP del instrumento
        self.ip_adress_entry = ttk.Entry(self.conection_frame)
        self.ip_adress_entry.grid(row=0, column=0, padx=5, pady=10)

        # Creación del botón para conectar el dispositivo
        self.button_probe = ttk.Button(self.conection_frame, text='Probar conexión', command=self.probe_conection)
        self.button_probe.grid(row=0, column=1, padx=5, pady=10)

        # Creación del label de estado de la conexión
        self.conection_text = tk.StringVar()
        self.conection_text.set("Esperando conexión")
        self.label_status = ttk.Label(self.conection_frame, textvariable=self.conection_text)
        self.label_status.grid(row=0, column=2, padx=5, pady=5)


        ## Segundo Frame
        # Creación del label de selección de puerto
        self.label_port = ttk.Label(self.parameters_frame, text='Seleccione el puerto:')
        self.label_port.grid(row=0, column=0, padx=2, pady=2, sticky='w')

        # Creación de los botones para seleccionar puerto de entrada
        self.port = tk.IntVar()
        self.port_button = ttk.Radiobutton(self.parameters_frame, text='50Ω', variable=self.port, value=50)
        self.port_button.grid(row=0, column=1, padx=2, pady=2)
        self.port_button = ttk.Radiobutton(self.parameters_frame, text='75Ω', variable=self.port, value=75)
        self.port_button.grid(row=0, column=2, padx=2, pady=2)

        # Creación del label de selección de transductores
        self.label_port = ttk.Label(self.parameters_frame, text='Seleccione los transductores:')
        self.label_port.grid(row=1, column=0, padx=2, pady=2, sticky='w')

        # Nombre de las bandas que disponibles para la medida
        name_transducers = ['TELEVES','HL223','HL223 V1','CABLE']

        # Creación del diccionario necesario para crear los check buttons
        self.transducers = dict()
        for name in name_transducers:
            self.transducers[name] = tk.BooleanVar()

        # Creación de los check buttons
        i = 0
        for key, var in self.transducers.items():
            check_button = ttk.Checkbutton(self.parameters_frame, text=key, variable=var)
            check_button.grid(row=i//3 + 1, column=i%3 + 1, padx=10, pady=2, sticky='w')
            i += 1

        ## Tercer Frame
        # Creación del label de instrucción para ingreso de número de canales
        self.num_of_channels_label = ttk.Label(self.channels_frame, text='Ingrese el número de canales que se medirán')
        self.num_of_channels_label.grid(row=0, column=0, padx=5, pady=2)

        # Creación del Entry para el ingreso de número de canales
        self.num_of_channels_entry = ttk.Entry(self.channels_frame, width=5)
        self.num_of_channels_entry.grid(row=0, column=1, padx=5, pady=2)

        # Creación del botón para conectar el dispositivo
        self.add_channels_button = ttk.Button(self.channels_frame, text='Crear', command=self.create_list)
        self.add_channels_button.grid(row=0, column=2, padx=5, pady=2)

        ## Cuarto Frame
        # Creación del botón para iniciar la medida    
        self.button_start = ttk.Button(self.measurement_frame, text='Iniciar medición', command=self.start_measurement)
        self.button_start.grid(row=0, column=0, padx=10, pady=10)

        # Creación del label de estado de la conexión
        self.measure_text = tk.StringVar()
        self.measure_text.set('Esperando medición')
        self.label_measure = ttk.Label(self.measurement_frame, textvariable=self.measure_text)
        self.label_measure.grid(row=0, column=1, padx=10, pady=10)

        # Obtener la altura requerida para los nuevos widgets
        self.window.update_idletasks()
        new_width  = self.window.winfo_reqwidth()
        new_height = self.window.winfo_reqheight()

        self.window.geometry(f"{new_width}x{new_height}")

    # Función del botón 'Probar conexión'
    def probe_conection(self):
        # Obtención de la IP ingresada por el usuario
        # En caso de estar vacía, la rutina buscará los dispositivos disponibles
        self.ip_adress = self.ip_adress_entry.get()

        # Conexión con el dispositivo
        try:
            self.etl = DTV(self.ip_adress, True, False)
            self.etl.instrument_status_checking = True  # Error check after each command
            self.etl.data_chunk_size = 100 # Definición del tamaño del buffer
            idn = self.etl.idn_string.split(sep=',')[1]
            self.conection_text.set(f'Conectado a: {idn}')
        
        # En caso de no lograrse, se le notifica al usuario en el Label
        except:
            self.conection_text.set('Error en la conexión')

        # Aplicación de reset al dispositivo
        self.etl.clear_status()
        self.etl.reset()

    # Función del botón 'Crear'
    def create_list(self):
        number_of_channels = self.num_of_channels_entry.get()
        try:
            number_of_channels = int(number_of_channels)
        except ValueError:
            messagebox.showerror(message='El valor ingresado no es un número')
        
        self.channels_and_plps = []
        for i in range(number_of_channels):
            channel_label = ttk.Label(self.channels_frame, text=f'Ingrese el canal # {i+1}')
            channel_label.grid(row=i+1, column=0, padx=5, pady=2, sticky='w')
            channel_entry = ttk.Entry(self.channels_frame, width=5)
            channel_entry.grid(row=i+1, column=1, padx=5, pady=2)
            plp_label = ttk.Label(self.channels_frame, text=f'# de PLPs')
            plp_label.grid(row=i+1, column=2, padx=5, pady=2, sticky='w')
            plp_entry = ttk.Entry(self.channels_frame, width=5)
            plp_entry.grid(row=i+1, column=3, padx=5, pady=2, sticky='w')
            self.channels_and_plps.append([channel_entry, plp_entry])

        # Obtener la altura requerida para los nuevos widgets
        self.window.update_idletasks()
        new_width  = self.window.winfo_reqwidth()
        new_height = self.window.winfo_reqheight()

        self.window.geometry(f"{new_width}x{new_height}")

    # Función del botón 'Iniciar medición'
    def start_measurement(self):

        # Creación de carpeta de soportes
        self.etl.create_folder()

        # Obtención del puerto de entrada de la antena
        input = self.port.get()

        # Obtención de los transductores
        transducers = []

        # Se crea el listado de bandas a medir, según los check button
        for text, var in self.transducers.items():
            if var.get() == True:
                transducers.append(text)

        # Se inicia la medición de cada canal
        for channel, number_of_plp in self.channels_and_plps:
            
            channel = channel.get()
            number_of_plp = number_of_plp.get()

            try:
                channel = int(channel)
                number_of_plp = int(number_of_plp)
            except ValueError:
                messagebox.showerror(message='Uno de los valores ingresados no es un número')

            # Mensaje de que banda se está leyendo en el Label
            self.measure_text.set(f'Midiendo el canal {channel}')
            
            # Mensaje de aviso para apuntar la antena hacia el acimuth deseado
            messagebox.showinfo(message=f'Apunte la antena hacia el acimuth para el canal {channel} y luego de clic en aceptar')

            # Puesta en marcha de la medida del canal
            dtv_results = self.etl.measurement(input, transducers, channel, number_of_plp)

            # Llenado del excel
            self.etl.fill_excel_channel_sheets(channel, dtv_results)

        # Mensaje de medición finalizada en el Label
        self.measure_text.set('Medición finalizada')

        # Cierra conexión con el instrumento
        self.etl.close()


# Ejecución del programa
if __name__ == '__main__':
    app = WelcomeWindow()