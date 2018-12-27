import pymysql.cursors
from openpyxl import Workbook
import re
# Connect to the database
connection = pymysql.connect(host='192.168.100.39',
                             user='thaonv',
                             password='meditech2017',
                             db='nagios',
                             charset='utf8mb4',
                             cursorclass=pymysql.cursors.DictCursor)

# workbook = Workbook()
# excelFilename = 'Bao_cao_chi_tieu_KPI_server.xlsx'
# sheet1 = workbook.active
# sheet1.title = "range names"
# sheet1.cell(1, 1, 'IP address')
# sheet1.cell(1, 2, 'Availability min')
# sheet1.cell(1, 3, 'Availability max')
# sheet1.cell(1, 4, 'Availability average')
# sheet1.cell(1, 5, 'CPU Utilization min')
# sheet1.cell(1, 6, 'CPU Utilization max')
# sheet1.cell(1, 7, 'CPU Utilization average')
# sheet1.cell(1, 8, 'Memory Used min')
# sheet1.cell(1, 9, 'Memory Used max')
# sheet1.cell(1, 10, 'Memory Used average')
# sheet1.cell(1, 11, 'Giảm trừ hợp lệ')
# sheet1.cell(1, 12, 'Thời gian downtime')
try:
    with connection.cursor() as cursor:
        sql = "SELECT start_time, end_time,    output FROM nagios_hostchecks WHERE host_object_id=265"
        cursor.execute(sql)
        print(cursor.fetchall())
        print(cursor.rowcount)
    # # Host
    # with connection.cursor() as cursor:
    #     sql = "SELECT host_object_id, display_name, address FROM nagios_hosts"
    #     cursor.execute(sql)
    #     hostGroup = cursor.fetchall()
    #
    # # CPU
    # with connection.cursor() as cursor:
    #     sql = "SELECT host_object_id, service_object_id FROM nagios_services WHERE display_name='CPU utilization'"
    #     cursor.execute(sql)
    #     cpuGroup = cursor.fetchall()
    #
    # # memory
    # with connection.cursor() as cursor:
    #     sql = "SELECT host_object_id, service_object_id FROM nagios_services WHERE display_name IN ('Memory', 'Memory used')"
    #     cursor.execute(sql)
    #     memoryGroup = cursor.fetchall()
    #
    # for hostIndex, hostItem in enumerate(hostGroup):
    #     sheet1.cell(hostIndex+2, 1, hostItem['address'])
    #     for cpuItem in cpuGroup:
    #         cpuObjectId = None
    #         if cpuItem['host_object_id'] == hostItem['host_object_id']:
    #             cpuObjectId = cpuItem['service_object_id']
    #         if cpuObjectId is not None:
    #             with connection.cursor() as cursor:
    #                 sql = "SELECT output FROM nagios_servicechecks WHERE service_object_id='{}'".format(cpuObjectId)
    #                 cursor.execute(sql)
    #                 cpuDetail = {}
    #                 cpuSum = float(0)
    #                 cpuNumberItem = cursor.rowcount
    #                 for cursorIndex, cursorItem in enumerate(cursor.fetchall()):
    #                     if 'total' in cursorItem['output']:
    #                         cpuUsed = float(re.findall(r"total: ([\d.]*)%", cursorItem['output'])[0])
    #                     else:
    #                         cpuUsed = float(re.findall(r"OK - ([\d.]*)%", cursorItem['output'])[0])
    #                     if cursorIndex == 0:
    #                         cpuDetail['min'] = cpuUsed
    #                         cpuDetail['max'] = cpuUsed
    #                     else:
    #                         cpuDetail['min'] = min(cpuUsed, cpuDetail['min'])
    #                         cpuDetail['max'] = max(cpuUsed, cpuDetail['max'])
    #                         cpuSum += cpuUsed
    #                 cpuDetail['min'] = round(cpuDetail['min'], 2)
    #                 cpuDetail['max'] = round(cpuDetail['max'], 2)
    #                 cpuDetail['average'] = round((cpuSum / cpuNumberItem), 2)
    #
    #                 sheet1.cell(hostIndex + 2, 5, cpuDetail['min'])
    #                 sheet1.cell(hostIndex + 2, 6, cpuDetail['max'])
    #                 sheet1.cell(hostIndex + 2, 7, cpuDetail['average'])
    #
    #     for memoryItem in memoryGroup:
    #         memoryObjectId = None
    #         if memoryItem['host_object_id'] == hostItem['host_object_id']:
    #             memoryObjectId = memoryItem['service_object_id']
    #         if memoryObjectId is not None:
    #             with connection.cursor() as cursor:
    #                 sql = "SELECT output FROM nagios_servicechecks WHERE service_object_id='{}'".format(memoryObjectId)
    #                 cursor.execute(sql)
    #                 memoryDetail = {}
    #                 memorySum = float(0)
    #                 memoryNumberItem = cursor.rowcount
    #                 for cursorIndex, cursorItem in enumerate(cursor.fetchall()):
    #                     memoryUsed = re.findall(r"OK - ([\d.]*)%", cursorItem['output'])
    #                     if len(memoryUsed) == 0:
    #                         if 'RAM used' in cursorItem['output']:
    #                             memoryUsed = float(re.findall(r"GB \(([\d.]*)%\)", cursorItem['output'])[0])
    #                         else:
    #                             break
    #                     else:
    #                         memoryUsed = float(memoryUsed[0])
    #                     if cursorIndex == 0:
    #                         memoryDetail['min'] = memoryUsed
    #                         memoryDetail['max'] = memoryUsed
    #                     else:
    #                         memoryDetail['min'] = min(memoryUsed, memoryDetail['min'])
    #                         memoryDetail['max'] = max(memoryUsed, memoryDetail['max'])
    #                         memorySum += memoryUsed
    #                 memoryDetail['min'] = round(memoryDetail['min'], 2)
    #                 memoryDetail['max'] = round(memoryDetail['max'], 2)
    #                 memoryDetail['average'] = round((memorySum / memoryNumberItem), 2)
    #                 sheet1.cell(hostIndex + 2, 8, memoryDetail['min'])
    #                 sheet1.cell(hostIndex + 2, 9, memoryDetail['max'])
    #                 sheet1.cell(hostIndex + 2, 10, memoryDetail['average'])
    #
    # workbook.save(filename=excelFilename)
finally:
    connection.close()
