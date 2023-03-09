import sys
from PyQt5.QtWidgets import QApplication, QWidget, QComboBox, QLabel, QLineEdit, QPushButton, \
                            QDateEdit,QHBoxLayout, QVBoxLayout
from PyQt5.QtCore import Qt, QDate, QTime
from PyQt5.QtGui import QIcon
from Google import Create_Service

class GoogleDataEntry(QWidget): 
    CLIENT_SECRET_FILE = 'client_secrets.json'
    API_NAME = 'sheets'
    API_VERSION = 'v4'
    SCOPES = ['https://www.googleapis.com/auth/spreadsheets']
    TAB_NAME= "TAB NAME"
    GOOGLE_SPREADSHEET_ID='SHEETS ID'
    
    def __init__(self): 
        super().__init__()
        self.setWindowTitle('Form') 
        self.setWindowIcon(QIcon('sheets.ico'))
        self.setMinimumWidth(1100) 
        self.service = None
        
        self.layout = QVBoxLayout() 
        self.setLayout(self.layout)
        
        self.initUI()
        self.reset_fields()
        self.connect_to_sheets()  
        
    def connect_to_sheets(self):
        try:
            self.service = Create_Service(self.CLIENT_SECRET_FILE, self.API_NAME, self.API_VERSION, self.SCOPES) 
            self.update_status('Connected to Google Sheets', '#266953')
        except Exception as e:
            self.update_status(str(e), 'red')
            
    def update_status(self, text, color='#FFFFFF'):
        try:
            self.status.setText(text)
            self.status.setStyleSheet('color: {0};'.format(color))
        except:
            return       
            
    def add_entry(self): 
        if not self.service:
            self.update_status('Google Sheets service is not connected')
            return
        
        record = [
            self.lineEdit['COLUMN NAME 1'].text(), 
            self.lineEdit['COLUMN NAME 2'].text(),
            self.lineEdit['COLUMN NAME 3'].text(),
            self.lineEdit['COLUMN NAME 4'].text() 
        ]
        try:
            self.service.spreadsheets().values().append(
                spreadsheetId=self.GOOGLE_SPREADSHEET_ID,
                range=f'{self.TAB_NAME}!A2',
                valueInputOption='USER_ENTERED', 
                insertDataOption='INSERT_ROWS',
                body={
                'majorDimension': 'ROWS', 
                'values': [record]
            }
        ).execute()
            self.reset_fields()
            self.update_status('Entry added','#266953')
        
        except Exception as e:
            self.update_status(str(e), 'red')
 
 
 
    def initUI(self):
        self.lineEdit = {}
        self.lineEdit['COLUMN NAME 1'] = QLineEdit()
        self.lineEdit['COLUMN NAME 2'] = QLineEdit() 
        self.lineEdit['COLUMN NAME 3']= QDateEdit()
        self.lineEdit['COLUMN NAME 3'].setCalendarPopup(True)
        self.lineEdit['COLUN NAME 4'] = QLineEdit() 

        layouts = {}
        layouts [1] = QHBoxLayout() 
        layouts [2]= QHBoxLayout()
        layouts [3]= QHBoxLayout()
        layouts [4] = QHBoxLayout()

        for k, v in layouts.items():
            self.layout.addLayout(v)
 
 
 
        layouts[1].addWidget(QLabel('Column 1: '), 2, alignment = Qt.AlignRight) 
        layouts [1].addWidget(self.lineEdit['COLUMN NAME 1'], 5)
        layouts[1].addWidget(QLabel('Column 2: '), 2, alignment = Qt.AlignRight) 
        layouts [1].addWidget(self.lineEdit['COLUMN NAME 2'], 5)
        layouts [2].addWidget(QLabel('Column 3'), 2, alignment = Qt.AlignRight) 
        layouts [2].addWidget(self.lineEdit['COLUMN NAME 3'], 4)
    
        layouts [2].addWidget(QLabel('Column 4'), 2, alignment = Qt.AlignRight) 
        layouts [2].addWidget(self.lineEdit['COLUMN NAME 4'], 4)

        layouts ['buttons'] = QHBoxLayout() 
        layouts['buttons'].addStretch() 
        self.layout.addLayout(layouts ['buttons'])
        buttons = {}
        buttons['Add'] = QPushButton('&Add',clicked = self.add_entry)
        buttons['Reset'] = QPushButton('&Reset', clicked = self.reset_fields)
        buttons['Close'] = QPushButton('&Close', clicked = app.quit)
        buttons['Close'].setStyleSheet('''
            color: red;
            width: 200px;
        ''')
        
        layouts['buttons'].addWidget(buttons['Add'])
        layouts['buttons'].addWidget(buttons['Reset'])
        layouts['buttons'].addWidget(buttons['Close'])
        self.status = QLabel()
        self.status.setStyleSheet('font-size: 25px')
        self.layout.addWidget(self.status)

    def reset_fields(self):
        self.lineEdit['COLUMN NAME 1'].setDate(QDate.currentDate()) 
        self.lineEdit['COLUMN NAME 2'].clear()
        self.lineEdit['COLUMN NAME 3'].clear()
        self.lineEdit['COLUMN NAME 4'].clear()     
        
        


if __name__ == '__main__':
# don't auto scale when drag app to a different monitor.
# QApplication.setAttribute(Qt.HighDpiScaleFactor RoundingPolicy. PassThrough)
        app = QApplication(sys.argv)
        app.setStyleSheet('''
            QWidget {
                font-size: 30px;
           }
            QPushButton {
                width:250px;
                height:45px;    
            }
            QLineEdit {
                height:45px;
            }
        ''')
        data_entry_form = GoogleDataEntry()
        data_entry_form.show()

        try:
            sys.exit(app.exec_())
        except SystemExit:
            print('Closing Window ...')
    