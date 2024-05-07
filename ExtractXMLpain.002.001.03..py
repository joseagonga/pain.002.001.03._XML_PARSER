import os
import pandas as pd
from lxml import etree
from collections import defaultdict

folder_path =  "C:\XML"
data = defaultdict(list)
data = []
data2 = defaultdict(list)
data2 = []
data3 = defaultdict(list)
data3 = []
data4 = defaultdict(list)
data4 = []
data5 = defaultdict(list)
data5 = []


for filename in os.listdir(folder_path):
  if filename.endswith('.xml'):
    xml_file = os.path.join(folder_path, filename)
    tree = etree.parse(xml_file)
    root = tree.getroot()
    transaction = {}
    
    #Almacenar en variables los valores unicos de fichero
    for tx_info in root.findall('.//{urn:iso:std:iso:20022:tech:xsd:pain.002.001.03}GrpHdr'):  
      CreDtTm = tx_info.find('.//{urn:iso:std:iso:20022:tech:xsd:pain.002.001.03}CreDtTm').text
      MsgId = tx_info.find('.//{urn:iso:std:iso:20022:tech:xsd:pain.002.001.03}MsgId').text

    for tx_info in root.findall('.//{urn:iso:std:iso:20022:tech:xsd:pain.002.001.03}InitgPty'):  
      Nm = tx_info.find('.//{urn:iso:std:iso:20022:tech:xsd:pain.002.001.03}Nm').text

    for tx_info in root.findall('.//{urn:iso:std:iso:20022:tech:xsd:pain.002.001.03}Othr'):  
      Id = tx_info.find('.//{urn:iso:std:iso:20022:tech:xsd:pain.002.001.03}Id').text     

    for tx_info in root.findall('.//{urn:iso:std:iso:20022:tech:xsd:pain.002.001.03}OrgnIGrpInfAndSts'):  
      OrgnlMsgId = tx_info.find('.//{urn:iso:std:iso:20022:tech:xsd:pain.002.001.03}OrgnlMsgId').text     
      OrgnIMsgNmId = tx_info.find('.//{urn:iso:std:iso:20022:tech:xsd:pain.002.001.03}OrgnIMsgNmId').text   
      OrgnINbOfTxs = tx_info.find('.//{urn:iso:std:iso:20022:tech:xsd:pain.002.001.03}OrgnINbOfTxs').text   
      OrgnICtrlSum = tx_info.find('.//{urn:iso:std:iso:20022:tech:xsd:pain.002.001.03}OrgnICtrlSum').text   


    for tx_info in root.findall('.//{urn:iso:std:iso:20022:tech:xsd:pain.002.001.03}Othr'):  
      Id = tx_info.find('.//{urn:iso:std:iso:20022:tech:xsd:pain.002.001.03}Id').text     




  # Extract data using XPath expressions
    for tx_info in root.findall('.//{urn:iso:std:iso:20022:tech:xsd:pain.002.001.03}TxInfAndSts'):
      transaction = {}
      transaction['MsgId'] = MsgId
      transaction['CreDtTm'] = CreDtTm
      transaction['Nm'] = Nm
      transaction['Id'] = Id
      transaction['NombreFichero'] = filename
      #transaction['OrgnIMsgld'] = OrgnlMsgId      
      #transaction['OrgnIMsgNmId'] = OrgnIMsgNmId 
      #transaction['OrgnINbOfTxs'] = OrgnINbOfTxs 
      #transaction['OrgnICtrlSum'] = OrgnICtrlSum
      #transaction['OrgnlMsgId'] = OrgnlMsgId
      #transaction['OrgnlMsgId'] = OrgnlMsgId        
      transaction['OrgnlEndToEndId'] = tx_info.find('.//{urn:iso:std:iso:20022:tech:xsd:pain.002.001.03}OrgnlEndToEndId').text
      transaction['TxSts'] = tx_info.find('.//{urn:iso:std:iso:20022:tech:xsd:pain.002.001.03}TxSts').text
      data.append(transaction)

    for tx_info in root.findall('.//{urn:iso:std:iso:20022:tech:xsd:pain.002.001.03}OrgnlTxRef'):
      transaction2 = {}
      transaction2['ReqdColltnDt'] = tx_info.find('.//{urn:iso:std:iso:20022:tech:xsd:pain.002.001.03}ReqdColltnDt').text
    
      data2.append(transaction2)

    for tx_info in root.findall('.//{urn:iso:std:iso:20022:tech:xsd:pain.002.001.03}Amt'):
      transaction3 = {}
      transaction3['InstdAmt'] = tx_info.find('.//{urn:iso:std:iso:20022:tech:xsd:pain.002.001.03}InstdAmt').text
    
      data3.append(transaction3)
 
    for tx_info in root.findall('.//{urn:iso:std:iso:20022:tech:xsd:pain.002.001.03}CdtrAcct'):
      transaction4 = {}
      transaction4['IBAN'] = tx_info.find('.//{urn:iso:std:iso:20022:tech:xsd:pain.002.001.03}IBAN').text
    
      data4.append(transaction4)

    for tx_info in root.findall('.//{urn:iso:std:iso:20022:tech:xsd:pain.002.001.03}DbtrAcct'):
      transaction5 = {}
      transaction5['IBAN2'] = tx_info.find('.//{urn:iso:std:iso:20022:tech:xsd:pain.002.001.03}IBAN').text
    
      data5.append(transaction5)    
    

df = pd.DataFrame(data)
df2 = pd.DataFrame(data2)
df3 = pd.DataFrame(data3)
df4 = pd.DataFrame(data4)
df5 = pd.DataFrame(data5)
df = pd.merge(left=df, right=df2, left_index=True, right_index=True, how='inner')
df = pd.merge(left=df, right=df3, left_index=True, right_index=True, how='inner')
df = pd.merge(left=df, right=df4, left_index=True, right_index=True, how='inner')
df = pd.merge(left=df, right=df5, left_index=True, right_index=True, how='inner')


#Copiar hasta aqui para print(df) de salida en PowerBI

pathCompletaOut = os.path.join(folder_path, "XML.xlsx")
df.to_excel(pathCompletaOut,index=False)