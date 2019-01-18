import sys
import os
from PyQt5.QtWidgets import *
import scikit_posthocs as ph
import ast
import openpyxl
import numpy
import pandas

def calculate():
    global combobox_posthoc_methods
    global list_of_lineedits
    global lineedit_output_file
    global a

    if(a is not None):
        ## Get function info
        method = getattr(ph, combobox_posthoc_methods.currentText())
        arguments = [a]
        for i in range(1, len(list_of_lineedits)):
            print(list_of_lineedits[i].text())
            try:
                arguments.append(ast.literal_eval(list_of_lineedits[i].text()))
            except ValueError:
                arguments.append(list_of_lineedits[i].text())

        print(arguments)
        result = method(*arguments)
        print(result)
        outframe = pandas.DataFrame(result)
        outframe.to_excel(lineedit_output_file.text())

def getInputFile():
    global label_input_file
    global lineedit_output_file
    global a

    fileName, _ = QFileDialog.getOpenFileName(None, "QFileDialog.getOpenFileName()", "",
                                              "Excel Files (*.xlsx);;All Files (*)")
    label_input_file.setText(fileName)
    lineedit_output_file.setText(os.path.splitext(fileName)[0]+'_output.xlsx')

    pandas_a = pandas.read_excel(fileName)
    a = pandas_a.T.values
    print(a)


def removeAllFromLayout(layout):
    for i in reversed(range(layout.count())):
        layout.itemAt(i).widget().setParent(None)

def combobox_posthoc_methods_selection_changed(text):
    global combobox_posthoc_methods
    global frame_method_arguments
    global layout_main
    global layout_labels
    global layout_values
    global list_of_labels
    global list_of_lineedits
    global a


    ## Get function info
    print('Selection changed ' + text)
    method = getattr(ph, text)
    argcount = method.__code__.co_argcount
    varnames = method.__code__.co_varnames
    argnames = varnames[0:argcount]
    defaults = method.__defaults__
    print(argnames)
    print(argcount)
    print(defaults)
    print(method.__kwdefaults__)

    #Set tooltip
    combobox_posthoc_methods.setToolTip(method.__doc__)

    #removeAllFromLayout(layout_labels)
    ## Create frame
    if len(list_of_labels)>0:
        for label in list_of_labels:
            layout_labels.removeWidget(label)
            label.setParent(None)
    if len(list_of_lineedits)>0:
        for lineedit in list_of_lineedits:
            layout_values.removeWidget(lineedit)
            lineedit.setParent(None)

    default_counter = 0
    list_of_lineedits = []
    for name in argnames:
        label = QLabel(name)
        list_of_labels.append(label)
        layout_labels.addWidget(label)

        if(name == 'x' or name == 'g'):
            lineedit = QLineEdit('NOT IMPLEMENTED YET')
            lineedit.setEnabled(False)
        elif(name == 'a'):
            if a is None:
                lineedit = QLineEdit('Not loaded yet')
                lineedit.setEnabled(False)
            else:
                lineedit = QLineEdit('Loaded from file')
                lineedit.setEnabled(False)
        else:
            lineedit = QLineEdit(str(defaults[default_counter]))
            default_counter = default_counter+1

        list_of_lineedits.append(lineedit)
        layout_values.addWidget(lineedit)

    #for default in defaults:
    #    lineedit = QLineEdit(str(default))
    #    list_of_lineedits.append(lineedit)
    #    layout_values.addWidget(lineedit)
    #    #print(ast.literal_eval(str(default)))



if __name__ == '__main__':
    global app
    global combobox_posthoc_methods
    global frame_method_arguments
    global layout_main
    global list_of_labels
    global list_of_lineedits
    global label_input_file
    global lineedit_output_file
    global a

    a = None

    app = QApplication(sys.argv)

    # Post Hoc Statistics Combo Box
    combobox_posthoc_methods = QComboBox()
    temp = dir(ph)
    sub  = str('posthoc')
    posthoc_methods_list = [s for s in temp if sub in s]
    for s in posthoc_methods_list:
        combobox_posthoc_methods.addItem(s)

    # Vertical Box Layout
    layout_main = QVBoxLayout()

    # Argument List Layout
    layout_arguments = QHBoxLayout()
    layout_labels    = QVBoxLayout()
    layout_values    = QVBoxLayout()

    # File layout
    layout_files = QVBoxLayout()
    layout_input_file        = QHBoxLayout()
    button_input_file_choose = QPushButton('Choose')
    button_input_file_choose.clicked.connect(getInputFile)
    label_input_file         = QLabel('')
    label_input_file_description = QLabel('Input file')
    layout_input_file.addWidget(label_input_file_description)
    layout_input_file.addWidget(label_input_file)
    layout_input_file.addWidget(button_input_file_choose)
    label_output_file_description = QLabel('Output file')
    layout_output_file   = QHBoxLayout()
    lineedit_output_file = QLineEdit('')
    layout_output_file.addWidget(label_output_file_description)
    layout_output_file.addWidget(lineedit_output_file)
    layout_files.addLayout(layout_input_file)
    layout_files.addLayout(layout_output_file)

    # Lower Buttons
    layout_buttons   = QHBoxLayout()
    button_calculate = QPushButton('Calculate')
    button_calculate.clicked.connect(calculate)
    layout_buttons.addWidget(button_calculate)

    layout_main.addLayout(layout_files)
    layout_main.addWidget(combobox_posthoc_methods)
    layout_arguments.addLayout(layout_labels)
    layout_arguments.addLayout(layout_values)
    layout_main.addLayout(layout_arguments)
    layout_main.addLayout(layout_buttons)
    list_of_labels    = []
    list_of_lineedits = []

    w = QWidget()
    w.resize(250, 150)
    w.move(300, 300)
    w.setWindowTitle('Statistic Electricity (by Murilo)')
    w.setLayout(layout_main)

    combobox_posthoc_methods.currentTextChanged.connect(combobox_posthoc_methods_selection_changed)

    w.show()

    sys.exit(app.exec_())
