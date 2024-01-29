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

    format_chart_data_wb(prs, fdata, 55, 3, '咖啡市场情况', 'F:G', 430, 11)
    format_chart_data_wb(prs, fdata, 55, 5, '咖啡市场情况', 'B:D', 430, 11)
    format_table_data_wb(prs, fdata, 55, 4, '咖啡市场情况', 'G:J', 430, 11, 1, -1)
    format_remark_text(prs, 55, 2, '生意参谋')
    prs.save(test_output)

if __name__ == '__main__':
    project_folder=os.path.dirname(os.path.dirname(os.path.realpath(__file__)))
    prs_data_load = os.path.join(project_folder,'./data/nespresso_rpt_data_202401.xlsx')
    prs_load = os.path.join(project_folder,'./data/nespresso_rpt_demo.pptx')
    test_output = os.path.join(project_folder,'./data/nespresso_rpt_test.pptx')

    test_page(prs_data_load, prs_load, test_output)