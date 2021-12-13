import os
import win32com.client as client
import PySimpleGUI as sg
import win32gui, win32con

The_program_to_hide = win32gui.GetForegroundWindow()
win32gui.ShowWindow(The_program_to_hide , win32con.SW_HIDE)

sg.theme('DarkBlack1')

layout = [
    [sg.Text('ORIENTAÇÕES')],
    [sg.Text('- O arquivo deve estar fechado para a execução da conversão')],
    [sg.Text('- Os arquivos deverão ser salvos na pasta Arquivos para que dê certo')],
    [sg.Text('')],
    [sg.Button('Converter')],
    [sg.Text('Desenvolvido por Rafael Lins Fontes')]
]
janela=sg.Window('Conversor xls para xlsx').layout(layout)

botão = janela.read()

excel = client.Dispatch("excel.application")

for file in os.listdir(os.getcwd() + "/arquivos/"):
    filename, fileextension = os.path.splitext(file)
    wb = excel.Workbooks.Open(os.getcwd() + "/arquivos/" + file)
    output_path = os.getcwd() + "/convertidos/" + filename
    wb.SaveAs(output_path,51)
    wb.Close()
excel.Quit()