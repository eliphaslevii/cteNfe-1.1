import sys,json,os
import json
from PyQt5.QtGui import QIcon
from PyQt5.QtWidgets import (QApplication, QFileDialog, QWidget, QFormLayout,
                             QLineEdit,QPushButton, QMessageBox,QVBoxLayout,QMainWindow,
                             QLabel)
import mysql.connector

def resourcePath(relativePath):
    try:
        basePath = sys._MEIPASS
    except AttributeError:
        basePath = os.path.dirname(__file__)
    return os.path.join(basePath,relativePath)

CONFIG_FILE = resourcePath('./config/config.json')

def load_config():
    if os.path.exists(CONFIG_FILE):
        with open(CONFIG_FILE, 'r') as file:
            return json.load(file)
    else:
        return {}
def save_config(config):
    with open(CONFIG_FILE,'w') as file:
        json.dump(config,file,indent=4)
class ConfigWindow(QWidget):
    def __init__(self):
        super().__init__()
        self.setWindowTitle('Configurações')
        self.setWindowIcon(QIcon('../xlsxLog/physics.png'))
        self.setGeometry(200, 200, 400, 300)
        self.loadConfig()
        self.init_ui()

    def init_ui(self):

        formLayout = QFormLayout()

        # Criando labels e campos de entrada
        self.dbUsernameLabel = QLabel("Usuário do Banco de Dados:")
        self.dbUsernameEdit = QLineEdit(self)
        self.dbUsernameEdit.setText(self.dbUsername)
        formLayout.addRow(self.dbUsernameLabel, self.dbUsernameEdit)

        self.dbPasswordLabel = QLabel("Senha do Banco de Dados:")
        self.dbPasswordEdit = QLineEdit(self)
        self.dbPasswordEdit.setEchoMode(QLineEdit.Password)
        self.dbPasswordEdit.setText(self.dbPassword)
        formLayout.addRow(self.dbPasswordLabel, self.dbPasswordEdit)

        self.databaseNameLabel = QLabel("Nome do Banco de Dados:")
        self.databaseNameEdit = QLineEdit(self)
        self.databaseNameEdit.setText(self.databaseName)
        formLayout.addRow(self.databaseNameLabel, self.databaseNameEdit)

        self.dbHostLabel = QLabel("Host do Banco de Dados:")
        self.dbHostEdit = QLineEdit(self)
        self.dbHostEdit.setText(self.dbHost)
        formLayout.addRow(self.dbHostLabel, self.dbHostEdit)

        self.testButton = QPushButton("Testar Conexão",self)
        self.testButton.clicked.connect(self.testDbConn)
        formLayout.addRow(self.testButton)

        saveButton = QPushButton("Salvar", self)
        saveButton.clicked.connect(self.saveConfig)
        formLayout.addRow(saveButton)

        self.setLayout(formLayout)

    def selectCertPath(self):
        file_name, _ = QFileDialog.getOpenFileName(self, 'Selecione o Arquivo PFX', '','Arquivos PFX (*.pfx);;Todos os Arquivos (*)')
        if file_name:
            self.certPathEdit.setText(file_name)
    def loadConfig(self):
        configFile = resourcePath('./config/config.json')
        if os.path.exists(configFile):
            with open(configFile,'r') as file:
                self.config = json.load(file)
                self.dbUsername = self.config.get('dbUsername')
                self.dbPassword = self.config.get('dbPassword')
                self.databaseName = self.config.get('databaseName')
                self.dbHost = self.config.get('dbHost')
                #self.certPath = self.config.get('certPath')
                #self.certPasswd = self.config.get('certPasswd')
        else:
            self.config = {}
            QMessageBox(self,"Warning","Arquivo de configuração não encontrado")

    def saveConfig(self):
        config = {
            "dbUsername":self.dbUsernameEdit.text(),
            "dbPassword":self.dbPasswordEdit.text(),
            "databaseName":self.databaseNameEdit.text(),
            "dbHost":self.dbHostEdit.text(),
            #"certPath":self.certPathEdit.text(),
            #"certPasswd":self.certPasswdEdit.text()
        }
        try:
            save_config(config)
            QMessageBox.information(self,"Success",'Configurações salvas com sucesso!')
            self.loadConfig()
        except Exception as e:
            QMessageBox.critical(self,'Error',f'Erro ao salvar configurações: {e}')

    def testDbConn(self):
        try:
            connection = mysql.connector.connect(
                host=self.dbHost,
                user=self.dbUsername,
                password=self.dbPassword,
                database=self.databaseName
            )
            if connection.is_connected():
                QMessageBox.information(None, "Success", "Conexão ao banco de dados bem-sucedida!")
                connection.close()
        except Exception as e:
                QMessageBox.critical(None, "Error", f"Erro ao conectar ao banco de dados: {e}")

if __name__ == '__main__':
    app = QApplication(sys.argv)
    main_window = ConfigWindow()
    main_window.show()
    sys.exit(app.exec_())
