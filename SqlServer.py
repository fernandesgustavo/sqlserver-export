import sys
import struct
import socket
import pymssql
import xlwt 
from xlwt import Workbook
from datetime import date

class SqlServer:

    def __init__(self, server, user, passwd, db):
        
        self.__server = server
        self.__user = user
        self.__passwd = passwd
        self.__db = db
        self.__connection = False

    def convertIp(self, ip_decimal):
        return socket.inet_ntoa(struct.pack("!L", ip_decimal)) # return ip address format

    def connect(self):
        try:
            self.__connection = pymssql.connect(self.__server, self.__user, self.__passwd, self.__db)
        except Exception as exception:
            print('Failed to connect DB: '+ str(exception))

    def read(self, query):
        cursor = self.__connection.cursor(as_dict=True)
        # cursor = self.__connection.cursor()
        cursor.execute(query)
        return cursor
        # cursor.encode('utf-8')

    def exportHardware(self):
        query = 'SELECT H.nId, H.wstrWinName, H.nIP, H.wstrComment, I.wstrManufacturer, I.strSerial, I.nCpuCores, I.nCpuThreads, I.nCapacity, I.wstrName, I.strMAC, H.nPlatformType FROM v_akpub_host AS H INNER JOIN v_akpub_hwinv AS I ON I.nHost = H.nId ORDER BY H.wstrWinName'
        today = date.today()
        self.connect()
        cursor = self.read(query)

        wb = Workbook()
        
        sheet1 = wb.add_sheet('hardware')

        style = xlwt.easyxf('font: bold 1')

        sheet1.write(0, 0, 'nId', style) 
        sheet1.write(0, 1, 'wstrWinName', style)
        sheet1.write(0, 2, 'nIP', style)
        sheet1.write(0, 3, 'wstrComment', style)
        sheet1.write(0, 4, 'wstrManufacturer', style)
        sheet1.write(0, 5, 'strSerial', style)
        sheet1.write(0, 6, 'nCpuCores', style)
        sheet1.write(0, 7, 'nCpuThreads', style)
        sheet1.write(0, 8, 'nCapacity', style)
        sheet1.write(0, 9, 'wstrName',style)
        sheet1.write(0, 10, 'strMAC', style)

        line = 1
        
        for row in cursor:
            sheet1.write(line, 0, row['nId'])
            sheet1.write(line, 1, row['wstrWinName'])
            sheet1.write(line, 2, self.convertIp(row['nIP']))
            sheet1.write(line, 3, row['wstrComment'])
            sheet1.write(line, 4, row['wstrManufacturer'])
            sheet1.write(line, 5, row['strSerial'])
            sheet1.write(line, 6, row['nCpuCores'])
            sheet1.write(line, 7, row['nCpuThreads'])
            sheet1.write(line, 8, row['nCapacity'])
            sheet1.write(line, 9, row['wstrName'])
            sheet1.write(line, 10, row['strMAC'])
            
            line += 1
        
        wb.save('hardware-{}.xls'.format(today))
    
    def exportSoftware(self):
        query = 'SELECT H.wstrWinName, H.nIP, A.wstrDisplayName, A.wstrBuild, A.wstrPublisher, A.tmInstallDate FROM v_akpub_host AS H INNER JOIN v_akpub_application AS A ON A.nHostId = H.nId ORDER BY H.wstrWinName'
        today = date.today()
        self.connect()
        cursor = self.read(query)

        wb = Workbook()
        
        sheet1 = wb.add_sheet('software')

        style = xlwt.easyxf('font: bold 1')

        sheet1.write(0, 0, 'wstrWinName', style) 
        sheet1.write(0, 1, 'nIP', style)
        sheet1.write(0, 2, 'wstrDisplayName', style)
        sheet1.write(0, 3, 'wstrBuild', style)
        sheet1.write(0, 4, 'wstrPublisher', style)
        sheet1.write(0, 5, 'tmInstallDate', style)

        line = 1
        
        for row in cursor:
            sheet1.write(line, 0, row['wstrWinName'])
            sheet1.write(line, 1, self.convertIp(row['nIP']))
            sheet1.write(line, 2, row['wstrDisplayName'])
            sheet1.write(line, 3, row['wstrBuild'])
            sheet1.write(line, 4, row['wstrPublisher'])
            sheet1.write(line, 5, row['tmInstallDate'])
            
            line += 1
        
        wb.save('software-{}.xls'.format(today))

if __name__ == '__main__':

    server = sys.argv[1]
    user = sys.argv[2]
    passwd = sys.argv[3]
    db = sys.argv[4]
    report = sys.argv[5]

    sqlserver = SqlServer(server, user, passwd, db)

    if report == 'hardware':
        sqlserver.exportHardware()
        print('Planilha de Hardware criada com sucesso.')
    elif report == 'software':
        sqlserver.exportSoftware()
        print('Planilha de Software criada com sucesso.')