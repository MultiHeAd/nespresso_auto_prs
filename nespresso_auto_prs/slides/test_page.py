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
    format_chart_data3(prs, fdata, 12, 3, 'dat_nespresso_mkt_rk_brand', '销售金额', ['销售金额_Top1品牌_占比', '销售金额_Top2品牌_占比', '销售金额_Top3品牌_占比', '销售金额_TOP4-10品牌_占比', '销售金额_TOP11-20品牌_占比', '销售金额_Top品牌_其他商家_占比'], 1,'咖啡机')
    prs.save(test_output)

if __name__ == '__main__':
    project_folder=os.path.dirname(os.path.dirname(os.path.realpath(__file__)))
    prs_data_load = os.path.join(project_folder,'./data/nespresso_rpt_data_202401.xlsx')
    prs_load = os.path.join(project_folder,'./data/nespresso_rpt_demo.pptx')
    test_output = os.path.join(project_folder,'./data/nespresso_rpt_test.pptx')

    test_page(prs_data_load, prs_load, test_output)