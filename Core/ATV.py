from RsInstrument import *
import os

class ATV(RsInstrument):

    # Inicialización de la clase
    def __init__(self, ip_adress: str = '', id_query: bool = True, reset: bool = False, options: str = None, direct_session: object = None):
        if ip_adress == '':
            resource_name = RsInstrument.list_resources("?*")[0]
        else:
            resource_name = f'TCPIP::{ip_adress}::INSTR'
        super().__init__(resource_name, id_query, reset, options, direct_session) 
    
    # Nombre que se le dará a la carpeta de soportes
    folder_name = f'.\\Soportes_tv_analogica'

    # Contador para nombrar las carpetas, en caso de que ya esté creada
    count = 0

    # Definición del diccionario con las frecuencias centrales de los canales
    tvTable = {
        2:  57,   3: 63,   4: 69,   5: 79,   6: 85,   7: 177,  8: 183,
        9:  189, 10: 195, 11: 201, 12: 207, 13: 213, 14: 473, 15: 479,
        16: 485, 17: 491, 18: 497, 19: 503, 20: 509, 21: 515, 22: 521,
        23: 527, 24: 533, 25: 539, 26: 545, 27: 551, 28: 557, 29: 563,
        30: 569, 31: 575, 32: 581, 33: 587, 34: 593, 35: 599, 36: 605,
        38: 617, 39: 623, 40: 629, 41: 635, 42: 641, 43: 647, 44: 653,
        45: 659, 46: 665, 47: 671, 48: 677, 49: 683, 50: 689, 51: 695
    }

    # Función de configuración inicial del dispositivo
    def setup(self):
        
        # Aplicación del reset al instrumento
        self.clear_status()
        self.reset()

        # Configuraciones generales para todos los canales
        self.write_str_with_opc(f'SYST:POS:GPS OFF')
        self.write_str_with_opc(f'DET RMS')
        self.write_str_with_opc(f'INP:ATT 0 dB')
        self.write_str_with_opc(f'INP:ATT:AUTO OFF')
        self.write_str_with_opc(f'INP:GAIN:STAT OFF')
        self.write_str_with_opc(f'UNIT:POW DBUV')
        self.write_str_with_opc(f'FREQ:SPAN 6.5 MHz')
        self.write_str_with_opc(f'DISP:TRAC:MODE AVERage')
        self.write_str_with_opc(f'INIT:CONT OFF')
        self.write_str_with_opc(f'SWE:COUN 10')
        self.write_str_with_opc(f'channel 10 kHz')
        self.write_str_with_opc(f'channel:VID 30 kHz')


    # Creación de la carpeta de soportes de la medida
    def create_folder(self):
        try:
            os.mkdir(self.folder_name)
        except FileExistsError:
            self.count += 1
            self.folder_name = f'Soportes_tv_analogica_{self.count}'
            self.create_folder()


    # Función de generación de soportes
    def get_supports(self,channel):

        # Definición de la dirección y nombre de los archivos en el instrumento
        img_path_instr = '\\Public\\Screen Shots\\SS.png'
        set_path_instr = '\\Public\\Datasets\\dataset.set'
        csv_path_instr = '\\Public\\Datasets\\datacsv.csv'

        # Generación de subcarpeta por cada canal
        try:
            os.mkdir(f'{self.folder_name}\\CH_{channel}')
            
            # Definición de la dirección y nombre de los archivos en el pc
            img_path_pc = f'.\\{self.folder_name}\\CH_{channel}\\CH_{channel}.png'
            set_path_pc = f'.\\{self.folder_name}\\CH_{channel}\\CH_{channel}.set'
            csv_path_pc = f'.\\{self.folder_name}\\CH_{channel}\\CH_{channel}.csv'
        except FileExistsError:
            os.mkdir(f'{self.folder_name}\\CH_{channel}_1')

            # Definición de la dirección y nombre de los archivos en el pc, en caso de que la carpeta ya exista
            img_path_pc = f'.\\{self.folder_name}\\CH_{channel}_1\\CH_{channel}.png'
            set_path_pc = f'.\\{self.folder_name}\\CH_{channel}_1\\CH_{channel}.set'
            csv_path_pc = f'.\\{self.folder_name}\\CH_{channel}_1\\CH_{channel}.csv'
        

        # Generación de la captura de pantalla *.png en el instrumento
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


    # Función que ejecuta la medición según la banda
    def measurement(self, channel):

        # Configuración del instrumento según el canal
        self.write_str_with_opc(f'FREQ:CENT {self.tvTable[channel]} MHz')
        self.write_str_with_opc(f'CALC:MARK1 ON')
        self.write_str_with_opc(f'CALC:MARK1:X {self.tvTable[channel] - 3} MHz')
        self.write_str_with_opc(f'CALC:MARK2 ON')
        self.write_str_with_opc(f'CALC:MARK2:X {self.tvTable[channel] - 1.75} MHz')
        self.write_str_with_opc(f'CALC:MARK3 ON')
        self.write_str_with_opc(f'CALC:MARK3:X {self.tvTable[channel] + 2.75} MHz')
        self.write_str_with_opc(f'CALC:MARK4 ON')
        self.write_str_with_opc(f'CALC:MARK4:X {self.tvTable[channel] + 3} MHz')
        self.write_str_with_opc(f'INIT;*WAI')
        
        # Generación y adquisición de soportes
        self.get_supports(channel)


if __name__ == '__main__':
    ip_adress = '192.168.1.1'
    try:
        fsh = ATV(ip_adress,True, True, "Simulate=True")
        fsh.instrument_status_checking = True  # Error check after each command
        fsh.data_chunk_size = 100
        print(f'Conectado a: {fsh.idn_string}')
    except Exception as ex:
        print('Error initializing the instrument session:\n' + ex.args[0])
        exit()
    
    fsh.setup()
    fsh.measurement(11)