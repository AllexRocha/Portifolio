import PyPDF2
import glob
import re
import mysql.connector
from mysql.connector import Error

#criar objeto com as informações da pasta de pdfs
arquivos = glob.glob(r"C:\Users\55619\Desktop\temp\pdfs\*.pdf")

#Criar tabela

try:
    #Criar conexão com o banco de dados
    conexao = mysql.connector.connect(host = 'localhost',database = 'random_data',user = 'root',password='123456')
    
    #criar tabela
    criar_tabela_SQL = """CREATE TABLE dados_vendas(
                          ID INT UNSIGNED NOT NULL AUTO_INCREMENT,
                          NOME_DESTINATARIO VARCHAR(30) NOT NULL,
                          ENDEREÇO VARCHAR(60) NOT NULL,
                          CEP VARCHAR(10) NOT NULL,
                          DATA_ENTREGA VARCHAR(15) NOT NULL,
                          NUMERO_OC INT UNSIGNED NOT NULL,
                          PRODUTO VARCHAR(200) NOT NULL,
                          PRIMARY KEY (ID)
                                                    );"""

    cursor = conexao.cursor()
    cursor.execute(criar_tabela_SQL)
    print("Tabela criada com sucesso !")
    
#caso dê erro
except mysql.connector.Error as erro:
    print("Falha ao criar tabela no MySQL: {}".format(erro))
    
    
# Pegar informações de cada pdf e salva no banco de dados MySQL 1 pdf por linha
for i in range(0, len(arquivos)):
    filename=re.sub(r".*\\", "", arquivos[i])
    
    arquivo = r"C:\Users\55619\Desktop\temp\pdfs" +"\\" + filename  
    lerPDF = PyPDF2.PdfFileReader(arquivo)
    pagina = lerPDF.getPage(0)
    conteudo = str(pagina.extractText())
   
    conteudo=re.sub(r"\n", "", conteudo)
    
    #separa informações usando regex
    nome = re.findall('(?<=Caro\(a\) ).[\w][\w\s]*',conteudo)[0]
    data = re.findall(r'\d{2}/\d{2}/\d{4}',conteudo)[0]
    produto = re.findall('(?<=produto ).*?(?= com)',conteudo)[0]
    numref = re.findall('\d{9}',conteudo)[0]
    endereco = re.findall('(?<=endereço ).*?(?= de CEP)',conteudo)[0]
    cep = re.findall(r'\d{5}-\d{3}',conteudo)[0]
    
     
    dado = '\'' + nome + '\',' + '\"' + endereco + '\",' + '\'' + cep + '\',' + '\'' + data + '\','  + str(numref) + ',' + '\"' + produto + '\"' + ');'
    
    inserir_dados = """INSERT INTO dados_vendas
                   (NOME_DESTINATARIO,ENDEREÇO,CEP,DATA_ENTREGA,NUMERO_OC,PRODUTO)
                   VALUES
                   ("""
    enviar = inserir_dados + dado
    
    conexao = mysql.connector.connect(host = 'localhost',database = 'random_data',user = 'root',password='123456')
    
    
    cursor = conexao.cursor()
    #Inserir registros no banco de dados MySQL
    
    cursor.execute(enviar)
    conexao.commit()
    
#fechar cursor e conexão com o MySQL    
cursor.close()
conexao.close()

