import os, sys
import threading
from typing import Text
import pytchat
import PySimpleGUI as sg
import re
import xlsxwriter
from time import sleep


class LoadWrite:

    def __init__(self):
        self.path = os.path.dirname(sys.argv[0])
        return self.__trigger_antivirus()
    
    def __trigger_antivirus(self):
        """Trigger antivirus as soon as possible, manipulating external files"""
        
        open(self.path + r'\test.docx', 'w')
        os.remove(self.path + r'\test.docx')

    
    def start_xlsx(self, saveas_path):
        """
        Create and open a new fresh xlsx file,
        set cell patterns, and
        return title writer function
        """

        self.saveas_path = saveas_path
        self.workbook = xlsxwriter.Workbook(self.saveas_path)
        self.worksheet = self.workbook.add_worksheet('Live Chat')
        self.worksheet.set_column(0, 0, 20)
        self.worksheet.set_column(1, 1, 60)
        self.fmt_title = self.workbook.add_format({
            'bold'       : True,
            'align'      : 'center',
            'font_color' : 'white',
            'bg_color'   : 'black',
        })
        self.fmt_msg = self.workbook.add_format({
            'text_wrap'    : True,
            'valign'       : 'top',
            'text_justlast': True,
            'top'          : 1,
            'bottom'       : 1,
        })

        return self.write_title()
    
    def write_title(self):
        """Write a stylish title"""
        
        titles = ['Usuário', 'Comentário']
        row = 0
        for col, content in enumerate(titles):
            self.worksheet.write(row, col, content, self.fmt_title)

        return

    def live_chat(self, video_id: str, window):
        """
        Create a YT live spec,
        prints the log for the active SGWindow, and
        write a new row as updates
        """

        print(f'Live ID: {video_id}\n')
        row = 0
        last_msg = ''
        while not self.stop_thread:
            chat = pytchat.create(video_id=video_id, interruptable=False)
            while chat.is_alive():
                if self.stop_thread == True:
                    break
                data = chat.get()
                for item in data.items:
                    if self.stop_thread == True:
                        break
                    time = item.datetime[10:-3]
                    user = item.author.name
                    msg  = item.message.strip()
                    text = f"{time} [{user}] - {msg}"
                    if text == last_msg:
                        continue
                    else:
                        last_msg = text
                        row  += 1
                        self.worksheet.write_row(row, 0, (user, msg), self.fmt_msg)
                        window['_COUNT_'].update(row)
                        print(text)
            print('Conexão perdida. Tentando novamente em 1 segundo...')
            sleep(1)


        print('\n' + '==='*5)
        print(f'Fim da transmissão.\n{row} mensagens coletadas')
        
        self.close_xlsx()

    def close_xlsx(self):
        print(f'Salvando em {self.saveas_path}')
        while True:
            try:
                self.workbook.close()
                break
            except Exception:
                print('[ERRO] Arquivo aberto em outro programa,\ntentando novamente em 5 segundos...\n')
                sleep(5)
        print('Arquivo Salvo!')
        print('==='*5)


class SGWindow(LoadWrite):

    def __init__(self):
        LoadWrite.__init__(self, )

        sg.set_options(use_ttk_buttons=True, ttk_theme='clam', enable_treeview_869_patch=True)
        sg.theme_background_color('#181818')
        sg.theme_text_color('#FFFFFF')
        sg.theme_input_background_color('#3785C8')
    
    def __layout_main(self):
        column_count = [
            [
                sg.Text('Mensagens coletadas:', text_color='#AAAAAA', background_color='#3D3D3D', pad=(8, 8)),
            ],
            [
                sg.Text('0', key='_COUNT_', background_color='#3D3D3D', font='Any 40', size=(4, 1), pad=((8, 26)), justification='center'),
            ],
        ]
        column_link = [
            [
                sg.Text('Link da live:', pad=(8, 0), text_color='#243441', background_color='#3785C8')
            ],
            [
                sg.Input('', font=('any 16'), text_color='#181818', key='_LINK_', size=(42, 1), pad=(8, (0, 4)), border_width=0),
            ],
        ]
        column_output = [
            [
                sg.Column(column_link, '#3785C8', pad=(0, 0)),
            ],
            [
                sg.Column(column_count, '#3D3D3D', pad=(0, 0)),
                sg.Output(size=(50, 8), pad=(0, 0), background_color='#3D3D3D', text_color='#FFFFFF'),
            ],
        ]
        column_buttons = [
            [
                sg.Button('COMEÇAR', key='_START_', button_color='#3785C8', size=(20, 1), pad=(12, (12, 6)), border_width=0),
            ],
            [
                sg.Button('PARAR E SALVAR', key='_STOP_', button_color='#181818', size=(20, 1), pad=(12, 6), border_width=0, disabled=True),
            ],
            # [
            #     sg.Checkbox('FORÇAR REPLAY', True, key='_REPLAY_', pad=(12, 14), background_color='#3D3D3D', tooltip='Coleta dados desde o início da live,\nmesmo se ainda estiver no ar')
            # ],
            [
                sg.Button(disabled=True, size=(20, 1), pad=(0, 12), border_width=0, button_color='#3D3D3D')
            ],
            [
                sg.Button('SALVAR COMO', key='_SAVE_AS_', button_type=3, target='_FILE_', file_types=(('Planilha do Microsoft Excel (*.xlsx)', '.xlsx'),), default_extension='.xlsx', size=(20, 1), pad=(8, (4, 0)), border_width=0, button_color='#243441'),
            ],
            [
                sg.Input('', key='_FILE_', size=(21, 1), pad=(0, (0, 12)), disabled=True),
            ],
        ]

        layout = [
            [
                sg.Column(column_buttons, '#3D3D3D', pad=((0, 12), 0), element_justification='center'),
                sg.Column(column_output, '#3785C8', pad=(0, 0)),
            ],
        ]

        return layout
    
    def window_main(self):

        def start(window, values):
            self.stop_thread = False
            link = values['_LINK_']
            path = values['_FILE_']
            if not link and not path:
                print('Ops! Faltou o link da live e um local para salvar o arquivo')
                return
            elif not link:
                print('Ops! Faltou o link da live!')
                return
            elif not path:
                print('Ops! Faltou um local para savar o arquivo!')
                return
            live_id = re.search('v=[\w-]+', link)
            if not live_id:
                print('Não foi possível identificar o ID da live')
                return
            live_id = live_id.group()[2:]

            LoadWrite.start_xlsx(self, path)
            self.thread = threading.Thread(target=LoadWrite.live_chat, args=(self, live_id, window))
            self.thread.daemon = True
            self.thread.start()

            window['_START_'].update(disabled=True, button_color='#181818')
            window['_STOP_'].update(disabled=False, button_color='#3785C8')
            # window['_REPLAY_'].update(disabled=True)
            window['_SAVE_AS_'].update(disabled=True, button_color='#181818')
        
        def stop(self, window):
            self.stop_thread = True
            window['_START_'].update(disabled=False, button_color='#3785C8')
            window['_STOP_'].update(disabled=True, button_color='#181818')
            # window['_REPLAY_'].update(disabled=False)
            window['_SAVE_AS_'].update(disabled=False, button_color='#243441')

            
        self.thread = threading.Thread()
        window = sg.Window('YT Live Chat', self.__layout_main(), margins=(12, 12), element_justification='center')
        while True:
            event, values = window.read(1000)

            # i += 1
            # window['_COUNT_'].update(f'{i}')
            if event == sg.WINDOW_CLOSED:
                break

            if event == '_START_':
                start(window, values)
            if event == '_STOP_':
                stop(self, window)
        self.stop_thread = True
        while self.thread.is_alive():
            sleep(0.5)

        window.close()


if __name__ == '__main__':
    SGWindow().window_main()
