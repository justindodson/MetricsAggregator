from PyQt5 import QtWidgets


def no_export_path_warning():
    msg = QtWidgets.QMessageBox()
    msg.setIcon(QtWidgets.QMessageBox.Warning)

    msg.setText('No File Path Selected')
    msg.setInformativeText('Please select a file path to export the document')
    msg.setStandardButtons(QtWidgets.QMessageBox.Ok)
    msg.exec_()
