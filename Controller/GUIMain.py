from View.metricsGUI import Ui_MainWindow
from PyQt5 import QtWidgets, QtCore
import sys


class WindowEXEC:
    REGIONS = {'Hub 1': 1, 'Hub 2': 2, 'Hub 3': 3, 'Hub 4': 4, 'Hub 5': 5, 'Canada': 6}

    def __init__(self):
        app = QtWidgets.QApplication(sys.argv)
        window = QtWidgets.QMainWindow()
        self.ui = Ui_MainWindow()
        self.ui.setupUi(window)

        self.get_region_selection()
        self.metrics_picker()
        self.trigger_file_selection_dialog()
        self.import_button_triggered()
        self.detect_list_selected()
        self.remove_button_triggered()
        self.trigger_file_save_dialog()
        self.export_button_triggered()

        window.show()
        sys.exit(app.exec_())

    def metrics_picker(self):
        self.ui.metricsSelector.currentTextChanged.connect(self.get_metric_type)

    def get_metric_type(self):
        return self.ui.metricsSelector.currentText()

    def get_region_selection(self):
        region = 0
        if self.ui.hub1.isChecked():
            region = 1
        elif self.ui.hub2.isChecked():
            region = 2
        elif self.ui.hub3.isChecked():
            region = 3
        elif self.ui.hub4.isChecked():
            region = 4
        elif self.ui.hub5.isChecked():
            region = 5
        elif self.ui.canada.isChecked():
            region = 6

        return region

    def trigger_file_selection_dialog(self):
        self.ui.fileDialogButton.clicked.connect(self.open_file_dialog)

    def open_file_dialog(self):
        dig = QtWidgets.QFileDialog()
        dig.setFileMode(QtWidgets.QFileDialog.AnyFile)

        if dig.exec_():
            filename = dig.selectedFiles()
            self.show_file_name(filename)

    def show_file_name(self, filepath):
        self.ui.fileSelectInput.setText(filepath[0])

    def import_button_triggered(self):
        self.ui.importFileButton.clicked.connect(self.import_file_path)

    def import_file_path(self):
        if self.ui.fileSelectInput.text() != '':
            self.ui.filePathList.addItem(QtWidgets.QListWidgetItem(self.ui.fileSelectInput.text()))
            self.ui.fileSelectInput.clear()
            self.ui.fileSelectInput.repaint()

    def detect_list_selected(self):
        self.ui.filePathList.itemClicked.connect(self.enable_remove_button)

    def enable_remove_button(self):
        if len(self.ui.filePathList.selectedItems()) > 0:
            self.ui.removeButton.setEnabled(True)

    def remove_button_triggered(self):
        self.ui.removeButton.clicked.connect(self.remove_file_from_list)

    def remove_file_from_list(self):
        items = self.ui.filePathList.selectedItems()
        for item in items:
            self.ui.filePathList.takeItem(self.ui.filePathList.row(item))

    def trigger_file_save_dialog(self):
        self.ui.fileExportDialog.clicked.connect(self.open_save_file_dialog)

    def open_save_file_dialog(self):
        dig = QtWidgets.QFileDialog()
        filename = dig.getSaveFileName()
        self.show_save_filepath(filename[0])

    def show_save_filepath(self, filepath):
        self.ui.fileExportPathInput.setText(filepath)
        self.ui.fileExportPathInput.repaint()

    def export_button_triggered(self):
        self.ui.exportButton.clicked.connect(self.export_aggregated_metrics_document)

    def export_aggregated_metrics_document(self):
        path = self.ui.fileExportPathInput.text()
        if path != '':
            # todo: Create the method to call the appropriate method.
            pass
        else:
            self.show_no_export_path_warning_dialog()

    def __show_no_export_path_warning_dialog(self):
        msg = QtWidgets.QMessageBox()
        msg.setIcon(QtWidgets.QMessageBox.Warning)

        msg.setText('No File Path Selected')
        msg.setInformativeText('Please select a file path to export the document')
        msg.setStandardButtons(QtWidgets.QMessageBox.Ok)
        msg.exec_()


if __name__ == "__main__":
    win = WindowEXEC()
