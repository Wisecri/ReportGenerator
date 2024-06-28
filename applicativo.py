# Author: Cristiano Pelliccia

from utils import *

def create_window(layout):
    return sg.Window('Report Generator', layout, return_keyboard_events=True,
                     location=(300, 150), use_default_focus=False, size=(800, 500),\
                     icon='images' + os.sep + 'icon.ico')

window = create_window(get_layout_initial_window())

while True:
    event, values = window.read()

    if event == sg.WIN_CLOSED or event == "Exit":
        break
    if event == 'REPORT_TYPE':
        selected_report_type = values['REPORT_TYPE']
        window.close()
        window = create_window(get_layout_window()[selected_report_type])

    elif event == 'Back':
        window.close()
        window = create_window(get_layout_initial_window())

    elif event == 'IN_FILE_1':
        file_1 = sg.popup_get_file('', no_window=True, icon='images' + os.sep + 'icon.ico',
                                   initial_folder=os.path.join(os.path.join(os.environ['USERPROFILE']), 'Desktop'))
        window['FILE_1'].update(file_1)

    elif event == 'IN_FILE_23':
        file_23 = sg.popup_get_file('', no_window=True, icon='images' + os.sep + 'icon.ico',
                                    initial_folder=os.path.join(os.path.join(os.environ['USERPROFILE']), 'Desktop'))
        window['FILE_23'].update(file_23)

    elif event == 'IN_FILE_POT':
        file_pot = sg.popup_get_file('', no_window=True, icon='images' + os.sep + 'icon.ico',
                                    initial_folder=os.path.join(os.path.join(os.environ['USERPROFILE']), 'Desktop'))
        window['FILE_POT'].update(file_pot)

    elif event == 'CREATE_CDLS':
        cdl = values['CDL']
        file_1 = values['FILE_1']
        file_23 = values['FILE_23']
        if cdl and file_1 and file_23:
            try:
                create_pdf_cdls(file_1, file_23, cdl)
                sg.Popup('Report created successfully!', keep_on_top=True, icon='images' + os.sep + 'icon.ico')
            except Exception as e:
                sg.Popup(f'Error: {e}', keep_on_top=True, icon='images' + os.sep + 'icon.ico')

    elif event == 'CREATE_DIPS':
        dip = values['DIP']
        file_1 = values['FILE_1']
        file_23 = values['FILE_23']
        if dip and file_1 and file_23:
            try:
                create_pdf_dips(file_1, file_23, dip)
                sg.Popup('Report created successfully!', keep_on_top=True, icon='images' + os.sep + 'icon.ico')
            except Exception as e:
                sg.Popup(f'Error: {e}', keep_on_top=True, icon='images' + os.sep + 'icon.ico')

    elif event == 'CREATE_POT_CDL':
        cdl_pot = values['CDL_POT']
        uni_pot = values['UNI_POT']
        file_pot = values['FILE_POT']
        if cdl_pot and uni_pot and file_pot:
            try:
                create_pdf_pot(file_pot, cdl_pot, uni_pot)
                sg.Popup('Report created successfully!', keep_on_top=True, icon='images' + os.sep + 'icon.ico')
            except Exception as e:
                sg.Popup(f'Error: {e}', keep_on_top=True, icon='images' + os.sep + 'icon.ico')

    elif event == 'CREATE_POT_UNI':
        uni_pot = values['UNI_POT']
        file_pot = values['FILE_POT']
        if uni_pot and file_pot:
            try:
                create_pdf_pot_uni(file_pot, uni_pot)
                sg.Popup('Report created successfully!', keep_on_top=True, icon='images' + os.sep + 'icon.ico')
            except Exception as e:
                sg.Popup(f'Error: {e}', keep_on_top=True, icon='images' + os.sep + 'icon.ico')

window.close()