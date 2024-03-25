from excel_form import ExcelForm
import sys
from PyQt5.QtWidgets import QApplication

app = QApplication(sys.argv)
app.setOrganizationName('Scloud')
app.setApplicationName('Monthly Sales Report Generator')
excel_form = ExcelForm()
excel_form.show()
sys.exit(app.exec_())
