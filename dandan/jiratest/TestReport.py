#!/usr/local/bin/python3
import matplotlib.pyplot as plt
import numpy as np
import os
import pandas as pd
import requests
import time
import xlrd
import xlwt
from PIL import Image
from jira import JIRA
from xlutils.copy import copy

author = 'huyaodan'


class Jira_Report(object):

    def __init__(self):
        # 创建一个workbook 设置编码
        self.workbook = xlwt.Workbook(encoding='utf-8')
        # 创建一个worksheet
        self.worksheet = self.workbook.add_sheet('My Worksheet')
        self.reportname = 'Test_Report_{}.xlsx'.format(time.strftime("%Y%m%d_%H%M%S", time.localtime()))
        # 写入列标签
        inf = zip(range(9),['Key', 'Summary', 'bug数', 'PM', 'DEV', 'case数', '通过数', '失败数', 'QA'])
        for col, label in inf:
            self.worksheet.write(0, col, label)
        self.R, self.issues = self.get_jira_issues()

    def get_jira_issues(self):
        # 从jira获取feature
        base_url = "https://project.atcloudbox.com/"
        jira = JIRA(base_url, basic_auth=('yaodan.hu@rdc-west.com', '123456abc!'))
        pro = input("产品名：")
        ver = input("版本号：")
        jql = "project = %s AND affectedVersion = %s AND issuetype = Feature45" % (pro, ver)

        if pro[0] == 'C':
            prokey = pro[0] + pro[3].upper() + pro[-3].upper() + pro[-1]
        else:
            prokey = pro[0] + pro[4].upper() + pro[-1]

        # 获取versionId
        version = jira.get_project_version_by_name(project=prokey, version_name=ver)
        verId = version.id

        # 获取测试循环
        headers = {
            'Content-Type': 'application/json',
        }
        params = (
            ('projectKey', prokey),
            ('versionId', verId),
            ("expand", "executionSummaries"),
        )
        response = requests.get('https://project.atcloudbox.com/rest/zapi/latest/cycle', headers=headers, params=params,
                                auth=('yaodan.hu@rdc-west.com', '123456abc!'))
        R = response.json()
        issues = jira.search_issues(jql)
        return R, issues

    # 获取测试循环里的用例数、失败数、创建人
    def cycle(self, key1):
        total = 0
        status = 0
        fail = 0
        Created = 0
        for key in self.R:
            if type(self.R[key]).__name__ == 'dict':
                Name = self.R[key]['name'].split(':')[0]
                if Name == key1:
                    total = self.R[key]['totalExecutions']
                    Created = self.R[key]['createdByDisplay']
                    for i in self.R[key]['executionSummaries']['executionSummary']:
                        if i['statusName'] == '通过':
                            status = i['count']
                        if i['statusName'] == '失败':
                            fail = i['count']
        return total, status, fail, Created

    def add_report(self):
        # 测试数据写入excel
        for i in self.issues:
            TotalExe, Pass, Fail, Created = self.cycle(i.key)

            # 参数对应 行, 列, 值
            info = zip(range(9), [i.key, i.fields.summary, len(i.fields.issuelinks), str(i.fields.reporter), str(i.fields.assignee),
                                  TotalExe, Pass, Fail, Created])
            for j, k in info:
                self.worksheet.write(self.issues.index(i) + 1, j, k)
        # 保存
        self.workbook.save(self.reportname)

    def dec_order_of_bug(self):
        # 将表格按照bug数降序排列
        scExcel = pd.read_excel(self.reportname)
        scExcel.sort_values(by='bug数', ascending=False, inplace=True)
        scExcel.to_excel('bug_top5.xlsx')
        BUG = scExcel['bug数']
        CASE = scExcel['case数']
        KEY = scExcel['Key']
        return BUG, CASE, KEY

    def get_bug_picture(self):
        # 报告top5输出成柱形图
        # 解决 plt 中文显示的问题
        BUG, CASE, KEY = self.dec_order_of_bug()
        plt.rcParams['font.sans-serif'] = ['Arial Unicode MS']
        Y = BUG[:5]
        Y1 = CASE[:5]
        X = KEY[:5]
        bar_width = 0.3  # 条形宽度
        index_male = np.arange(len(X))  # bug条形图的横坐标
        index_female = index_male + bar_width  # case条形图的横坐标
        plt.bar(index_male, height=Y, width=bar_width, color='r', label='bug数')
        plt.bar(index_female, height=Y1, width=bar_width, color='b', label='case数')

        plt.legend()  # 显示图例
        plt.xticks(index_male + bar_width / 2, X)  # 让横坐标轴刻度显示 Feature， index_male + bar_width/2 为横坐标轴刻度的位置
        plt.ylabel('数量')  # 纵坐标轴标题
        plt.xlabel('Feature')  # 横坐标轴标题
        plt.title('Feature-Bug数/Case数')  # 图形标题

        for a, b in zip(index_male, Y):  # 柱子上的数字显示
            plt.text(a, b, '%d' % b, ha='center', va='bottom', fontsize=12);
        for a, b in zip(index_male + bar_width, Y1):
            plt.text(a, b, '%d' % b, ha='center', va='bottom', fontsize=12);

        barchart = 'bug_top'
        plt.savefig(barchart + '.png', dpi=400)

        # 柱形图转bmp格式
        im = Image.open('%s' % os.path.join(os.getcwd(), barchart + '.png')).convert("RGB")
        im.save('%s.bmp' % barchart)
        return barchart

    def get_first_sheetobj(self):
        # 打开想要更改的excel文件
        # 将操作文件对象拷贝，变成可写的workbook对象
        old_excel = xlrd.open_workbook(self.reportname)
        new_excel = copy(old_excel)
        # 获得第一个sheet的对象
        barchart = self.get_bug_picture()
        ws = new_excel.get_sheet(0)
        ws.insert_bitmap(barchart + '.bmp', 30, 0, scale_x=0.2, scale_y=0.2)
        Newreportname = 'New_Test_Report_{}.xls'.format(time.strftime("%Y%m%d_%H%M%S", time.localtime()))
        new_excel.save(Newreportname)

if __name__ == '__main__':
    J = Jira_Report()
    J.add_report()
    J.get_first_sheetobj()
