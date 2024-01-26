from openpyxl import load_workbook
from pptx import Presentation
from pptx.chart.data import CategoryChartData
from pptx.dml.color import RGBColor
import datetime
import pandas as pd
import numpy as np
from pptx.util import Pt
from pptx.enum.text import PP_PARAGRAPH_ALIGNMENT
from generate_slides import *
import os

def test_page(prs_data_load, prs_load, test_output):
    fdata = prs_data_load
    prs = Presentation(prs_load)

    format_chart_data_sort(prs, fdata, 26, 4, 'dat_nespresso_shop_ct_chl', ['二级', '三级'], ['访客数_同比月', '访客数_本月', '支付转化率_同比月', '支付转化率_本月'], '广告流量', '咖啡机')
    col_no_pay=format_table_data_chl_sort(prs, fdata, 26, 3, 'dat_nespresso_shop_ct_chl', ['chl'], ['支付人数_同比', '访客数_同比', '支付转化率_同比'], '广告流量', '咖啡机')
    format_remark_text(prs, 26, 2, '生意参谋',  '生意参谋暂不支持查看{}渠道的后链路转化表现'.format('、'.join(col_no_pay)))

    prs.save(test_output)

if __name__ == '__main__':
    project_folder=os.path.dirname(os.path.dirname(os.path.realpath(__file__)))
    prs_data_load = os.path.join(project_folder,'./data/nespresso_rpt_data_202401.xlsx')
    prs_load = os.path.join(project_folder,'./data/nespresso_rpt_demo.pptx')
    test_output = os.path.join(project_folder,'./data/nespresso_rpt_test.pptx')

    test_page(prs_data_load, prs_load, test_output)