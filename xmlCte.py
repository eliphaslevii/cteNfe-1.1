import json, sys, os
import xml.etree.ElementTree as ET

from PyQt5.QtGui import QColor, QTextCursor, QTextCharFormat, QIcon
from PyQt5.QtWidgets import (QApplication, QMainWindow, QPushButton, QFileDialog,
                             QMessageBox, QProgressBar, QTableWidget, QTableWidgetItem,
                             QPlainTextEdit, QVBoxLayout, QWidget, QMenu, QAction)
import mysql.connector
from config import ConfigWindow


class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("Processamento de XMLs")
        self.setWindowIcon(QIcon('../xlsxLog/physics.png'))
        self.setGeometry(100, 100, 600, 600)
        self.red = (205, 5, 41)
        self.blue = (70, 120, 187)
        self.gray = QColor(197, 197, 197)
        self.initUI()
        ConfigWindow.loadConfig(self)
        self.connDbInstance()

    def initUI(self):

        layout = QVBoxLayout()
        self.centralWidgets = QWidget(self)
        self.setCentralWidget(self.centralWidgets)
        self.centralWidgets.setLayout(layout)

        self.menuBar = self.menuBar()
        self.configMenu = QMenu("&File", self)
        self.menuBar.addMenu(self.configMenu)

        self.configAction = QAction('Configurações', self)
        self.configAction.triggered.connect(self.openConfigWindow)
        self.configMenu.addAction(self.configAction)

        self.btnSelectCteFolder = QPushButton("Importar DACTE", self)
        self.btnSelectCteFolder.clicked.connect(self.selectCteFolder)

        self.btnSelectNfeFolder = QPushButton("Importar NFE", self)
        self.btnSelectNfeFolder.clicked.connect(self.selectNfeFolder)

        self.progressBar = QProgressBar(self)
        self.progressBar.setValue(0)

        self.consoleView = QPlainTextEdit()
        self.consoleView.setReadOnly(True)

        self.table = QTableWidget(self)
        self.table.setColumnCount(3)
        self.table.setHorizontalHeaderLabels(['Cte', 'Tag', 'Observação'])

        layout.addWidget(self.btnSelectCteFolder)
        layout.addWidget(self.btnSelectNfeFolder)
        layout.addWidget(self.progressBar)
        layout.addWidget(self.table)
        layout.addWidget(self.consoleView)

    def openConfigWindow(self):
        self.configWindow = ConfigWindow()
        self.configWindow.show()
    def displayMessage(self, stringMessage, color):
        cursor = self.consoleView.textCursor()
        cursor.movePosition(QTextCursor.End)  # Move o cursor para o fim do texto atual

        # Defina o formato do texto
        format = QTextCharFormat()
        format.setForeground(QColor(*color))  # Defina a cor do texto

        cursor.setCharFormat(format)
        cursor.insertText(stringMessage + '\n')  # Insira o texto com a formatação

        # Role a visão para mostrar a última mensagem
        self.consoleView.verticalScrollBar().setValue(self.consoleView.verticalScrollBar().maximum())
    def selectCteFolder(self):
        folder_path = QFileDialog.getExistingDirectory(self, "Selecionar Pasta")
        if folder_path:
            self.proccessCteFilesInFolder(folder_path)
    def selectNfeFolder(self):
        nfeFolderPath = QFileDialog.getExistingDirectory(self, "Selecionar Pasta")
        if nfeFolderPath:
            self.processNfeFilesInFolder(nfeFolderPath)
    def processNfeFilesInFolder(self, nfeFolderPath):
        fileNfePath = [os.path.join(nfeFolderPath, filename) for filename in os.listdir(nfeFolderPath) if
                       filename.endswith('.xml')]
        if not fileNfePath:
            QMessageBox.warning(self, "Atenção", "Nenhum arquivo XML encontrado na pasta selecionada.")
            return

        self.progressBar.setMaximum(len(fileNfePath))
        validNfeItems, invalidNfeItems = self.proccessNfeFiles(fileNfePath)
        self.saveNfeToDatabase(validNfeItems)
        self.poplateTable(validNfeItems, invalidNfeItems)
        QMessageBox.information(self, "Sucesso",
                                f"{len(validNfeItems)} arquivos XML processados e salvos com sucesso no banco de dados!")
    def proccessCteFilesInFolder(self, folder_path):
        file_paths = [os.path.join(folder_path, filename) for filename in os.listdir(folder_path) if
                      filename.endswith('.xml')]
        if not file_paths:
            QMessageBox.warning(self, "Atenção", "Nenhum arquivo XML encontrado na pasta selecionada.")
            return

        self.progressBar.setMaximum(len(file_paths))
        valid_items, invalid_items, keyItems = self.proccessXmlFiles(file_paths)
        self.saveToDatabase(valid_items)
        self.poplateTable(valid_items, invalid_items)
        self.saveKeyNfe(keyItems)
        QMessageBox.information(self, "Sucesso",
                                f"{len(valid_items)} arquivos XML processados e salvos com sucesso no banco de dados!")
    def proccessNfeFiles(self, fileNfePath):
        validNfeItems = []
        invalidNfeItems = []

        self.progressBar.setValue(0)
        for index, fileNfePath in enumerate(fileNfePath, start=1):
            try:
                parser = ET.XMLParser(encoding='UTF-8')
                tree = ET.parse(fileNfePath, parser=parser)
                root = tree.getroot()

                # Definindo o namespace
                ns = {'nfe': 'http://www.portalfiscal.inf.br/nfe'}

                fileNfeName = os.path.basename(fileNfePath)
                self.displayMessage(f'Info: Processando arquivo: {fileNfeName}', self.blue)

                ideElement = root.find('.//nfe:ide', namespaces=ns)
                if ideElement is None:
                    invalidNfeItems.append({'Item': fileNfeName, 'Tag': 'ideElement',
                                            'Observação': 'Não é um arquivo NFe válido (não contém elemento ide)'})
                    continue

                destElement = root.find('.//nfe:dest', namespaces=ns)
                if destElement is None:
                    invalidNfeItems.append({'Item': fileNfeName, 'Tag': 'destElement',
                                            'Observação': 'Não é um arquivo NFe válido (não contém elemento dest)'})
                    continue

                totalElement = root.find('.//nfe:total', namespaces=ns)
                if totalElement is None:
                    invalidNfeItems.append({'Item': fileNfeName, 'Tag': 'totalElement',
                                            'Observação': 'Campo totalElement ausente ou vazio em Total'})
                    continue

                transpElement = root.find('.//nfe:transp', namespaces=ns)
                if transpElement is None:
                    invalidNfeItems.append({'Item': fileNfeName, 'Tag': 'totalElement',
                                            'Observação': 'Campo totalElement ausente ou vazio em Total'})
                    continue

                natOpElement = ideElement.find('nfe:natOp', namespaces=ns)
                if natOpElement is None:
                    invalidNfeItems.append(
                        {'Item': fileNfeName, 'Tag': 'natOp', 'Observação': 'Campo natOp ausente ou vazio em IDE'})
                    continue

                nNfElement = ideElement.find('nfe:nNF', namespaces=ns)
                if nNfElement is None:
                    invalidNfeItems.append(
                        {'Item': fileNfeName, 'Tag': 'nNfElement', 'Observação': 'Campo nNF ausente ou vazio em IDE'})
                    continue

                dhEmiElement = ideElement.find('nfe:dhEmi', namespaces=ns)
                if dhEmiElement is None:
                    invalidNfeItems.append(
                        {'Item': fileNfeName, 'Tag': 'dhEmi', 'Observação': 'Campo dhEmi ausente ou vazio em IDE'})
                    continue

                cnpjDestElement = destElement.find('nfe:CNPJ', namespaces=ns)
                if cnpjDestElement is None:
                    cnpjDestElement = destElement.find('nfe:CPF', namespaces=ns)
                if cnpjDestElement is None:
                    invalidNfeItems.append({'Item': fileNfeName, 'Tag': 'cnpjDestElement',
                                            'Observação': 'Campo CNPJ/CPF ausente ou vazio em dest'})
                    continue

                xNomeDestElement = destElement.find('nfe:xNome', namespaces=ns)
                if xNomeDestElement is None:
                    invalidNfeItems.append(
                        {'Item': fileNfeName, 'Tag': 'xNomeDest', 'Observação': 'Campo xNome ausente ou vazio em dest'})
                    continue

                xMunDestElement = destElement.find('.//nfe:enderDest/nfe:xMun', namespaces=ns)
                if xMunDestElement is None:
                    invalidNfeItems.append(
                        {'Item': fileNfeName, 'Tag': 'xMunDest', 'Observação': 'Campo xMun ausente ou vazio em dest'})
                    continue

                ufDestElement = destElement.find('.//nfe:enderDest/nfe:UF', namespaces=ns)
                if ufDestElement is None:
                    invalidNfeItems.append(
                        {'Item': fileNfeName, 'Tag': 'ufDest', 'Observação': 'Campo UF ausente ou vazio em dest'})
                    continue

                vFreteElement = totalElement.find('.//nfe:ICMSTot/nfe:vFrete', namespaces=ns)
                if vFreteElement is None:
                    invalidNfeItems.append({'Item': fileNfeName, 'Tag': 'vFreteElement',
                                            'Observação': 'Campo vFrete ausente ou vazio em ICMSTot'})
                    continue

                vNfElement = totalElement.find('.//nfe:ICMSTot/nfe:vNF', namespaces=ns)
                if vNfElement is None:
                    invalidNfeItems.append({'Item': fileNfeName, 'Tag': 'vNfElement',
                                            'Observação': 'Campo vNfElement ausente ou vazio em ICMSTot'})
                    continue

                vProdElement = totalElement.find('.//nfe:ICMSTot/nfe:vProd',namespaces=ns)
                if vProdElement is None:
                    invalidNfeItems.append({'Item': fileNfeName,'Tag':'vProd','Observação':'Campo vProd ausente ou vazio.'})
                    continue

                volElement = transpElement.find('nfe:vol', namespaces=ns)
                if volElement is None:
                    invalidNfeItems.append({'Item': fileNfeName, 'Tag': 'volElement',
                                            'Observação': 'Campo volElement ausente ou vazio em Transp'})
                    continue

                qVolElement = volElement.find('nfe:qVol', namespaces=ns)
                if qVolElement is None:
                    invalidNfeItems.append({'Item': fileNfeName, 'Tag': 'volElement',
                                            'Observação': 'Campo volElement ausente ou vazio em Transp'})
                    continue

                pesoLElement = volElement.find('nfe:pesoL', namespaces=ns)
                if pesoLElement is None:
                    invalidNfeItems.append({'Item': fileNfeName, 'Tag': 'pesoLElement',
                                            'Observação': 'Campo pesoLElement ausente ou vazio em Transp'})
                    continue

                pesoBelement = volElement.find('nfe:pesoB', namespaces=ns)
                if pesoBelement is None:
                    invalidNfeItems.append({'Item': fileNfeName, 'Tag': 'pesoLElement',
                                            'Observação': 'Campo pesoLElement ausente ou vazio em Transp'})
                    continue

                protNFeElement = root.find('.//nfe:protNFe', namespaces=ns)
                if protNFeElement is None:
                    invalidNfeItems.append({'Item': fileNfeName, 'Tag': 'protNFeElement',
                                            'Observação': 'Campo protNFeElement ausente ou vazio'})
                    continue

                chNFeElement = protNFeElement.find('nfe:infProt/nfe:chNFe', namespaces=ns)
                if chNFeElement is None:
                    invalidNfeItems.append({'Item': fileNfeName, 'Tag': 'chNFeElement',
                                            'Observação': 'Campo chNFeElement ausente ou vazio em protNFe'})
                    continue


                chNFe = chNFeElement.text.strip()
                pesoB = pesoBelement.text.strip()
                pesoL = pesoLElement.text.strip()
                qVol = qVolElement.text.strip()
                vNF = vNfElement.text.strip()
                vFrete = vFreteElement.text.strip()
                ufDest = ufDestElement.text.strip()
                xMunDest = xMunDestElement.text.strip()
                xNomeDest = xNomeDestElement.text.strip()
                cnpjDest = cnpjDestElement.text.strip()
                dhEmiFull = dhEmiElement.text.strip()
                dhEmi = dhEmiFull.split('T')[0]
                nNF = nNfElement.text.strip()
                natOp = natOpElement.text.strip()
                vProd = vProdElement.text.strip()

                validNfeItems.append({
                    'natOp': natOp,
                    'nNF': nNF,
                    'dhEmi': dhEmi,
                    'cnpjDest': cnpjDest,
                    'xNomeDest': xNomeDest,
                    'xMunDest': xMunDest,
                    'ufDest': ufDest,
                    'vFrete': vFrete,
                    'vNF': vNF,
                    'qVol': qVol,
                    'pesoL': pesoL,
                    'pesoB': pesoB,
                    'chNFe': chNFe,
                    'vProd':vProd
                })

                self.progressBar.setValue(index)

            except Exception as e:
                self.displayMessage(f'Erro: {e}', self.blue)
                continue

        return validNfeItems, invalidNfeItems
    def proccessXmlFiles(self, file_paths):
        valid_items = []
        invalid_items = []
        keyItems = []

        self.progressBar.setValue(0)
        for index, file_path in enumerate(file_paths, start=1):
            try:
                self.progressBar.setValue(index)
                tree = ET.parse(file_path)
                root = tree.getroot()

                # Definindo o namespace
                ns = {'cte': 'http://www.portalfiscal.inf.br/cte'}
                filename = os.path.basename(file_path);
                self.displayMessage(f'Info: Processando arquivo: {filename}', self.blue)

                nCT_elem = root.find('.//cte:ide/cte:nCT', namespaces=ns)
                if nCT_elem is None:
                    invalid_items.append(
                        {'Item': filename, 'Tag': 'nCt', 'Observação': 'Não é um CTE válido, o arquivo será ignorado'})
                    continue

                nCT = nCT_elem.text.strip()

                if nCT is None:
                    invalid_items.append(
                        {'Item': filename, 'Tag': 'nCt', 'Observação': 'Não é um CTE válido, o arquivo será ignorado'})
                    continue

                chCte = root.find('.//cte:protCTe/cte:infProt/cte:chCTe', namespaces=ns).text.strip() if root.find(
                    './/cte:protCTe/cte:infProt/cte:chCTe', namespaces=ns) is not None else 'NULO'

                # Obtendo dhEmi
                dhEmi_full_elem = root.find('.//cte:ide/cte:dhEmi', namespaces=ns)
                if dhEmi_full_elem is not None:
                    dhEmi_full = dhEmi_full_elem.text.strip()
                    dhEmi = '-'.join(reversed(dhEmi_full.split('T')[0].split('-'))) if dhEmi_full else 'NULO'
                else:
                    dhEmi = 'NULO'

                emit = root.find('.//cte:emit', namespaces=ns)
                dest = root.find('.//cte:dest', namespaces=ns)
                rem = root.find('.//cte:rem', namespaces=ns)
                infCarga = root.find('.//cte:infCarga', namespaces=ns)
                vPrest = root.find('.//cte:vPrest', namespaces=ns)

                # CNPJ Emitente
                cnpjEmit_elem = emit.find('.//cte:CNPJ', namespaces=ns)
                cpfEmit_elem = emit.find('.//cte:CPF', namespaces=ns)

                if cnpjEmit_elem is not None:
                    cnpjEmit = cnpjEmit_elem.text.strip()
                elif cpfEmit_elem is not None:
                    cnpjEmit = cpfEmit_elem.text.strip()
                else:
                    cnpjEmit = 'NULO'

                # CNPJ Destinatário
                cnpjDest_elem = dest.find('.//cte:CNPJ', namespaces=ns)
                cpfDest_elem = dest.find('.//cte:CPF', namespaces=ns)
                if cnpjDest_elem is not None:
                    cnpjDest = cnpjDest_elem.text.strip()
                elif cpfDest_elem is not None:
                    cnpjDest = cpfDest_elem.text.strip()
                else:
                    cnpjDest = '00'
                if cnpjDest == '00':
                    invalid_items.append({'Item': 'cnpjDest', 'Tag': cnpjDest, 'Observação': 'cnpjDest é 00'})
                    continue

                nomeEmit = emit.find('.//cte:xNome', namespaces=ns).text.strip() if emit is not None and emit.find(
                    './/cte:xNome', namespaces=ns) is not None else '00'
                if nomeEmit == '00':
                    invalid_items.append({'Item': 'nomeEmit', 'Tag': nomeEmit, 'Observação': 'nomeEmit é 00'})
                    continue

                nomeDest = dest.find('.//cte:xNome', namespaces=ns).text.strip() if dest is not None and dest.find(
                    './/cte:xNome', namespaces=ns) is not None else '00'
                if nomeDest == '00':
                    invalid_items.append({'Item': 'nomeDest', 'Tag': nomeDest, 'Observação': 'nomeDest é 00'})
                    continue

                munDest = dest.find('.//cte:enderDest/cte:xMun',
                                    namespaces=ns).text.strip() if dest is not None and dest.find(
                    './/cte:enderDest/cte:xMun', namespaces=ns) is not None else '00'
                if munDest == '00':
                    invalid_items.append({'Item': 'munDest', 'Tag': munDest, 'Observação': 'munDest é 00'})
                    continue

                ufDest = dest.find('.//cte:enderDest/cte:UF',
                                   namespaces=ns).text.strip() if dest is not None and dest.find(
                    './/cte:enderDest/cte:UF', namespaces=ns) is not None else '00'
                if ufDest == '00':
                    invalid_items.append({'Item': 'ufDest', 'Tag': ufDest, 'Observação': 'ufDest é 00'})
                    continue

                vCarga = infCarga.find('.//cte:vCarga',
                                       namespaces=ns).text.strip() if infCarga is not None and infCarga.find(
                    './/cte:vCarga', namespaces=ns) is not None else '00'
                if vCarga == '00.00':
                    invalid_items.append({'Item': 'vCarga', 'Tag': vCarga, 'Observação': 'vCarga é 00.00'})
                    continue

                vtPrest = vPrest.find('.//cte:vTPrest',
                                      namespaces=ns).text.strip() if vPrest is not None and vPrest.find(
                    './/cte:vTPrest', namespaces=ns) is not None else '00'
                if vtPrest == '00.00':
                    invalid_items.append({'Item': 'vtPrest', 'Tag': vtPrest, 'Observação': 'vtPrest é 00.00'})
                    continue

                cnpjRem = rem.find('.//cte:CNPJ', namespaces=ns).text.strip()

                nomeRem = rem.find('.//cte:xNome', namespaces=ns).text.strip()

                infNFe_elems = root.findall('.//cte:infNFe/cte:chave', namespaces=ns)
                for infNFe_elem in infNFe_elems:
                    chaveNfe = infNFe_elem.text.strip()
                    keyItems.append({
                        'keyCte': chCte,
                        'keyNfe': chaveNfe,
                        'dhEmi': dhEmi
                    })

                valid_items.append({
                    'nCT': nCT,
                    'chCte': chCte,
                    'dhEmi': dhEmi,
                    'cnpjEmit': cnpjEmit,
                    'nomeEmit': nomeEmit,
                    'cnpjDest': cnpjDest,
                    'nomeDest': nomeDest,
                    'munDest': munDest,
                    'ufDest': ufDest,
                    'vCarga': vCarga,
                    'vtPrest': vtPrest,
                    'cnpjRem': cnpjRem,
                    'nomeRem': nomeRem
                })

            except Exception as e:
                self.displayMessage(f'Erro: {e}', self.blue)
                continue

        return valid_items, invalid_items, keyItems
    def poplateTable(self, valid_items, invalid_items):
        self.table.setRowCount(len(invalid_items))
        for row, item in enumerate(invalid_items):
            self.table.setItem(row, 0, QTableWidgetItem(item['Item']))
            self.table.setItem(row, 1, QTableWidgetItem(item['Tag']))
            self.table.setItem(row, 2, QTableWidgetItem(item['Observação']))
    def connDbInstance(self):
        try:
            self.conn = mysql.connector.connect(
                host=self.dbHost,
                user=self.dbUsername,
                password=self.dbPassword,
                database=self.databaseName
            )
        except Exception as e:
            self.displayMessage(f"Erro: erro ao conectar com o banco de dados {e}", self.red)
    def saveToDatabase(self, items):
        try:
            cursor = self.conn.cursor()
            cursor.execute('''
                CREATE TABLE IF NOT EXISTS cte (
                    id INT AUTO_INCREMENT PRIMARY KEY,
                    nCT VARCHAR(255),
                    chCte VARCHAR(255) UNIQUE,
                    dhEmi VARCHAR(255),
                    cnpjEmit VARCHAR(20),
                    nomeEmit VARCHAR(255),
                    cnpjDest VARCHAR(20),
                    nomeDest VARCHAR(255),
                    munDest VARCHAR(255),
                    ufDest VARCHAR(22),
                    vCarga DECIMAL(10, 2),
                    vtPrest DECIMAL(10, 2),
                    cnpjRem VARCHAR(20),
                    nomeRem VARCHAR(255),
                    dateUpdate TIMESTAMP DEFAULT CURRENT_TIMESTAMP
                )
            ''')
            insert_cte_query = '''
                INSERT INTO cte (
                    nCT, chCte, dhEmi, cnpjEmit, nomeEmit,
                    cnpjDest, nomeDest, munDest, ufDest, vCarga, vtPrest, cnpjRem, nomeRem
                ) VALUES (%s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s)
            '''

            self.progressBar.setValue(0)

            for index, item in enumerate(items, start=1):
                data_cte_tuple = (
                    item['nCT'], item['chCte'], item['dhEmi'], item['cnpjEmit'], item['nomeEmit'],
                    item['cnpjDest'], item['nomeDest'], item['munDest'], item['ufDest'],
                    float(item['vCarga']), float(item['vtPrest']), item['cnpjRem'], item['nomeRem']
                )
                try:
                    cursor.execute(insert_cte_query, data_cte_tuple)
                    self.conn.commit()
                    self.displayMessage(f'Info: Dados gravados com sucesso: {item['chCte']}', self.blue)

                except mysql.connector.IntegrityError as e:
                    self.displayMessage(f'Error: Erro ao inserir registro (chave duplicada): {item['chCte']}', self.red)
                    self.conn.rollback()
                except Exception as e:
                    self.displayMessage(f'Error: {e}', self.red)
                    self.conn.rollback()

                self.progressBar.setValue(index)

        except mysql.connector.Error as e:
            self.displayMessage(f'Error: {e}', self.blue)

        except Exception as e:
            self.displayMessage(f'Error: {e}', self.blue)

        finally:
            if 'conn' in locals() and self.conn.is_connected():
                self.progressBar.setValue(0)
                cursor.close()
                self.conn.close()
    def saveNfeToDatabase(self, items):
        try:
            cursor = self.conn.cursor()

            # Criar a tabela se não existir
            cursor.execute('''
                CREATE TABLE IF NOT EXISTS nfe (
                    id INT AUTO_INCREMENT PRIMARY KEY,
                    chNFe VARCHAR(255) UNIQUE,
                    pesoB FLOAT,
                    pesoL FLOAT,
                    qVol VARCHAR(255),
                    vNF DECIMAL(10, 2),
                    vFrete DECIMAL(10, 2),
                    ufDest VARCHAR(20),
                    xMunDest VARCHAR(255),
                    xNomeDest VARCHAR(255),
                    cnpjDest VARCHAR(255),
                    dhEmi VARCHAR(255),
                    nNF VARCHAR(255),
                    natOp VARCHAR(255),
                    vProd DECIMAL(10,2),
                    dateUpdate TIMESTAMP DEFAULT CURRENT_TIMESTAMP
                )
            ''')

            # Definir a consulta de inserção
            insert_query = '''
                INSERT INTO nfe (
                    chNFe, pesoB, pesoL, qVol, vNF,
                    vFrete, ufDest, xMunDest, xNomeDest, cnpjDest, dhEmi, nNF, natOp, vProd
                ) VALUES (%s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s)
            '''

            self.progressBar.setValue(0)

            # Inserir os dados
            for index, item in enumerate(items, start=1):
                data_tuple = (
                    item.get('chNFe'), item.get('pesoB'), item.get('pesoL'), item.get('qVol'),
                    float(item.get('vNF', 0.0)),
                    float(item.get('vFrete', 0.0)), item.get('ufDest'), item.get('xMunDest'), item.get('xNomeDest'),
                    item.get('cnpjDest'),
                    item.get('dhEmi'), item.get('nNF'), item.get('natOp'), item.get('vProd')
                )
                try:
                    cursor.execute(insert_query, data_tuple)
                    self.conn.commit()
                    self.displayMessage(f'Registro inserido com sucesso: {item.get("chNFe")}', self.blue)

                except mysql.connector.IntegrityError:
                    self.displayMessage(f'Erro ao inserir registro (chave duplicada): {item.get("chNFe")}', self.red)
                    self.conn.rollback()

                except mysql.connector.Error as e:
                    self.displayMessage(f'Erro ao inserir registro: {e}', self.red)
                    self.conn.rollback()

                self.progressBar.setValue(index)

        except mysql.connector.Error as e:
            self.displayMessage(f'Erro na conexão com o banco de dados: {e}', self.red)

        except Exception as e:
            self.displayMessage(f'Erro inesperado: {e}', self.blue)

        finally:
            if 'cursor' in locals():
                cursor.close()
            if 'conn' in locals() and self.conn.is_connected():
                self.conn.close()
            self.progressBar.setValue(0)
    def saveKeyNfe(self, items):
        try:

            cursor = self.conn.cursor()

            cursor.execute('''
                CREATE TABLE IF NOT EXISTS key_cte_nfe (
                    id INT AUTO_INCREMENT PRIMARY KEY,
                    keyCte VARCHAR(255),
                    keyNfe VARCHAR(255),
                    dhEmi VARCHAR(255),
                    dateUpdate TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
                    CONSTRAINT unique_key_cte_nfe UNIQUE (keyCte, keyNfe)
                )
            ''')

            insert_key_query = '''
                INSERT INTO key_cte_nfe (
                    keyCte, keyNfe, dhEmi
                ) VALUES (%s, %s, %s)
            '''
            self.progressBar.setValue(0)

            for index, item in enumerate(items, start=1):
                data_key_tuple = (
                    item['keyCte'], item['keyNfe'], item['dhEmi']
                )

                try:
                    cursor.execute(insert_key_query, data_key_tuple)
                    self.conn.commit()
                    self.displayMessage(f'Registro chave CTE-NFe inserido com sucesso: {item}', self.blue)

                except mysql.connector.IntegrityError as e:
                    self.displayMessage(f'Info: Erro ao inserir registro (chave duplicada): {item}', self.red)
                    self.conn.rollback()
                except Exception as e:
                    self.displayMessage(f'Error: {e}', self.blue)
                    self.conn.rollback()

                self.progressBar.setValue(index)
                self.displayMessage(f'Info: Processando chave Nfe/Cte: {item['keyCte']}', self.blue)

            self.displayMessage(f'Info:{len(items)} items processados', self.blue)

        except mysql.connector.Error as e:
            self.displayMessage(f'Error: {e}', self.blue)

        except Exception as e:
            self.displayMessage(f'Error: {e}', self.blue)

        finally:
            if 'conn' in locals() and self.conn.is_connected():
                cursor.close()
                self.conn.close()


def main():
    app = QApplication(sys.argv)
    window = MainWindow()
    window.show()
    sys.exit(app.exec_())


if __name__ == '__main__':
    main()
