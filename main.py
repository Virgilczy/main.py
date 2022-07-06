import PySimpleGUI as sg

from k1 import *
from k2 import *
from k3 import *
from k4 import *
from k5 import *

I1 = ''
I2 = ''
I3 = ''


def excel():
    global I1, I2, I3
    layout = [
        [sg.Text('请选择需要导入的学位课表格:(表头为专业、班级、学位课)')],
        [sg.In(I1, key='-In1-', size=(50, 100), enable_events=True),
         sg.FileBrowse(target='-In1-', file_types=(('All Files', '*.xls'), ('All Files', '*.xlsx')))],
        [sg.Text('请选择需要导入的必修课表格:(表头为课程编号、课程名称、学分数)')],
        [sg.In(I2, key='-In2-', size=(50, 100), enable_events=True),
         sg.FileBrowse(target='-In2-', file_types=(('All Files', '*.xls'), ('All Files', '*.xlsx')))],
        [sg.Text('请选择需要导入年级总表:')],
        [sg.In(I3, key='-In3-', size=(50, 100), enable_events=True),
         sg.FileBrowse(target='-In3-', file_types=(('All Files', '*.xls'), ('All Files', '*.xlsx')))],
        [sg.Text('请选择需要导出的形式:(导出前请关闭excel)')],
        [sg.Button('导出不及格学生名单', key='-k1-')],
        [sg.Button('导出必修课经补考获得学分', key='-k2-')],
        [sg.Button('导出学位课累计总分表格', key='-k3-')],
        [sg.Button('导出公选课累计总分表格', key='-k4-')]
    ]
    window = sg.Window('EXCEL的自动导入导出_陈章杨', layout)

    while True:
        event, values = window.read()
        if event is None:
            break
        if event == '-In1-':
            I1 = window['-In1-'].get()
        if event == '-In2-':
            I2 = window['-In2-'].get()
        if event == '-In3-':
            I3 = window['-In3-'].get()
        if event == '-k1-':
            f1 = window['-In3-'].get()
            layout1 = [
                [sg.Text('导出为:(导出前请关闭excel！)')],
                [sg.In('不及格名单.xls', key='-Out1-', size=(50, 100))],
                [sg.Text('导出需要时间,点击确认导出后请耐心等待！')],
                [sg.Button('确定导出', key='-Out11-'), sg.Button('返回主菜单', key='-Out12-')]
            ]
            window.close()
            window = sg.Window('导出选择', layout1)
            while True:
                event, values = window.read()
                if event is None:
                    break
                if event == '-Out11-':
                    f2 = window['-Out1-'].get()
                    k1(f1, f2)
                if event == '-Out12-':
                    window.close()
                    excel()

        if event == '-k2-':
            f1 = window['-In3-'].get()
            f3 = window['-In2-'].get()
            layout2 = [
                [sg.Text('导出为:(导出前请关闭excel！)')],
                [sg.In('必修课经补考获得学分表格.xls', key='-Out2-', size=(50, 100))],
                [sg.Text('导出需要时间,点击确认导出后请耐心等待！')],
                [sg.Button('确定导出', key='-Out21-'), sg.Button('返回主菜单', key='-Out22-')]
            ]
            window.close()
            window = sg.Window('导出选择', layout2)
            while True:
                event, values = window.read()
                if event is None:
                    break
                if event == '-Out21-':
                    f2 = window['-Out2-'].get()
                    k2(f1, f2)
                    k5(f2, f3)
                if event == "-Out22-":
                    window.close()
                    excel()

        if event == '-k3-':
            f1 = window['-In3-'].get()
            f3 = window['-In1-'].get()
            layout3 = [
                [sg.Text('导出为:(导出前请关闭excel！)')],
                [sg.In('学位课表格.xls', key='-Out3-', size=(50, 100))],
                [sg.Text('导出需要时间,点击确认导出后请耐心等待！')],
                [sg.Button('确定导出', key='-Out31-'), sg.Button('返回主菜单', key='-Out32-')]
            ]
            window.close()
            window = sg.Window('导出选择', layout3)
            while True:
                event, values = window.read()
                if event is None:
                    break
                if event == '-Out31-':
                    f2 = window['-Out3-'].get()
                    k3(f1, f2, f3)
                if event == '-Out32-':
                    window.close()
                    excel()

        if event == '-k4-':
            f1 = window['-In3-'].get()
            layout4 = [
                [sg.Text('导出为:(导出前请关闭excel！)')],
                [sg.In('公选课学分表格.xls', key='-Out4-', size=(50, 100))],
                [sg.Text('导出需要时间,点击确认导出后请耐心等待！')],
                [sg.Button('确定导出', key='-Out41-'), sg.Button('返回主菜单', key='-Out42-')]
            ]
            window.close()
            window = sg.Window('导出选择', layout4)
            while True:
                event, values = window.read()
                if event is None:
                    break
                if event == '-Out41-':
                    f2 = window['-Out4-'].get()
                    k4(f1, f2)
                if event == '-Out42-':
                    window.close()
                    excel()
    window.close()


if __name__ == "__main__":
    excel()
