from RsInstrument import *
import os
import time

class BancoDeMediciones(RsInstrument):

    # Inicialización de la clase
    def __init__(self, ip_adress: str = '', id_query: bool = True, reset: bool = False, options: str = None, direct_session: object = None):
        if ip_adress == '':
            resource_name = RsInstrument.list_resources("?*")[0]
        else:
            resource_name = f'TCPIP::{ip_adress}::INSTR'
        super().__init__(resource_name, id_query, reset, options, direct_session)

    # Nombre que se le dará a la carpeta de soportes
    folder_name = f'.\\Soportes_banco_de_mediciones'

    # Contador para nombrar las carpetas, en caso de que ya esté creada
    count = 0

    # Definición del diccionario con los parámetros de cada banda
    bands = {
        #Band      Fi      Ff      VBW   RBW  ref  tr. mode   unit 
        '700':    [703,    803,    100,  30, -20, 'AVERage', 'DBM'],
        '850':    [824,    894,    100,  30, -20, 'AVERage', 'DBM'],
        '1900':   [1850,   1990,   100,  30, -20, 'AVERage', 'DBM'],
        'AWS1':   [1755,   1780,   30,   10, -20, 'MAXHold', 'DBM'],
        'AWS2':   [2155,   2180,   30,   10, -20, 'MAXHold', 'DBM'],
        'AWS3':   [2170,   2200,   30,   10, -20, 'MAXHold', 'DBM'],
        '2.5':    [2500,   2690,   100,  30, -20, 'AVERage', 'DBM'],
        '900_1':  [894,    960,    30,   10, -60, 'MAXHold', 'DBM'],
        '900_2':  [900,    928,    30,   10, -60, 'AVERage', 'DBM'],
        '3500':   [3300,   3700,   100,  30, -40, 'MAXHold', 'DBM'],
        '2.4GHz': [2400,   2483.5, 30,   10, -60, 'MAXHold', 'DBM'],
        '5GHz':   [5180,   5825,   100,  30, -60, 'MAXHold', 'DBM'],
        '2300MHz':[2300,   2400,   100,  30, -60, 'AVERage', 'DBM'],
        'FM':     [88,     108,    30,   10,  96, 'AVERage', 'DBUV']
    }

    # Función de configuración inicial del dispositivo
    def setup(self):
        
        # Aplicación del reset al instrumento
        self.clear_status()
        self.reset()

        # Configuraciones generales para todas las bandas
        self.write_str_with_opc(f'DET RMS')
        self.write_str_with_opc(f'INP:ATT 0 dB')
        self.write_str_with_opc(f'INP:ATT:AUTO OFF')
        self.write_str_with_opc(f'INP:GAIN:STAT OFF')


    # Función de generación de soportes
    def get_supports(self,band):

        # Generación de subcarpeta por cada banda
        os.mkdir(f'{self.folder_name}\\{band}')

        # Definición de la dirección y nombre de los archivos en el instrumento
        img_path_instr = '\\Public\\Screen Shots\\SS.png'
        set_path_instr = '\\Public\\Datasets\\dataset.set'
        csv_path_instr = '\\Public\\Datasets\\datacsv.csv'
        
        # Definición de la dirección y nombre de los archivos en el pc
        img_path_pc = f'.\\{self.folder_name}\\{band}\\ScreenShot_{band}.png'
        set_path_pc = f'.\\{self.folder_name}\\{band}\\DatasetChannel_{band}.set'
        csv_path_pc = f'.\\{self.folder_name}\\{band}\\DatasetChannel_{band}.csv'

        # Generación de la captura de pantalla en el instrumento
        self.write_str_with_opc(f"MMEM:NAME '{img_path_instr}'") # Creación de la imagen en la memoria
        self.write_str_with_opc("HCOP:IMM") # Captura de la pantalla
        self.query_opc()

        # Generación del archivo *.set en el instrumento
        self.write_str_with_opc(f"MMEM:STOR:STAT 1,'{set_path_instr}'")
        self.query_opc()

        # Generación del archivo *.csv en el instrumento
        self.write_str_with_opc(f"MMEMory:STORe:CSV:STATe 1,'{csv_path_instr}'")
        self.query_opc()

        # Transferencia de los archivos generados al pc
        self.read_file_from_instrument_to_pc(img_path_instr, img_path_pc)
        self.read_file_from_instrument_to_pc(set_path_instr, set_path_pc)
        self.read_file_from_instrument_to_pc(csv_path_instr, csv_path_pc)

        # Eliminación de los archivos en el instrumento
        self.write_str_with_opc(f"MMEMory:DELete '{img_path_instr}'")
        self.write_str_with_opc(f"MMEMory:DELete '{set_path_instr}'")
        self.write_str_with_opc(f"MMEMory:DELete '{csv_path_instr}'")

    
    # Creación de la carpeta de soportes de la medida
    def create_folder(self):
        try:
            os.mkdir(self.folder_name)
        except FileExistsError:
            self.count += 1
            self.folder_name = f'Soportes_banco_de_mediciones_{self.count}'
            self.create_folder()


    # Función que ejecuta la medición según la banda
    def measurement(self, band):

        # Configuración del instrumento según la banda
        self.write_str_with_opc(f'FREQ:STAR {self.bands[band][0]} MHz') # Configuración de la frecuencia inicial
        self.write_str_with_opc(f'FREQ:STOP {self.bands[band][1]} MHz') # Configuración de la frecuencia final
        self.write_str_with_opc(f'BAND:VID {self.bands[band][2]} kHz') # Configuración del video bandwidth
        self.write_str_with_opc(f'BAND {self.bands[band][3]} kHz') # Configuración del resolution bandwidth
        self.write_str_with_opc(f'UNIT:POW {self.bands[band][6]}') # Configuración de la unidad
        self.write_str_with_opc(f'DISP:TRAC:Y:RLEV {self.bands[band][4]}') # Configuración del nivel de referencia
        self.write_str_with_opc(f'DISP:TRAC:MODE {self.bands[band][5]}') # Configuración del modo de traza
        if self.bands[band][5] == 'AVERage':
            self.write_str_with_opc(f'INIT:CONT OFF') # Apagado del modo de barrido continuo
            self.write_str_with_opc(f'SWE:COUN 10') # Configuración del número de trazas
            self.write_str_with_opc(f'INIT;*WAI') # Inicio del barrido y espera de que se complete el número de trazas
        elif self.bands[band][5] == 'MAXHold':
            wait = float(self.query('SWE:TIME?')) # Obtención del tiempo de un barrido
            self.write_str_with_opc(f'INIT:CONT ON') # Encendido del modo de barrido continuo
            self.write_str_with_opc(f'INIT') # Inicio del barrido
            time.sleep(wait*10) # Espera a que se complete el número de trazas

        # Generación y adquisición de soportes
        self.get_supports(band)

if __name__ == '__main__':
    ip_adress = '192.168.1.1'
    try:
        fsh = BancoDeMediciones(ip_adress,True, True, "Simulate=True")
        fsh.instrument_status_checking = True  # Error check after each command
        fsh.data_chunk_size = 100
        print(f'Conectado a: {fsh.idn_string}')
    except Exception as ex:
        print('Error initializing the instrument session:\n' + ex.args[0])
        exit()
    fsh.measurement('700')