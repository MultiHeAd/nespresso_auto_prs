# -*- coding:utf-8 -*-

from slides.generate_slides import run

prs_data_load = r'C:\Users\Yinfinity\Desktop\nespresso_auto_prs\data\nespresso_rpt_data_202401.xlsx'
prs_load = r'C:\Users\Yinfinity\Desktop\nespresso_auto_prs\data\nespresso_rpt_demo.pptx'
prs_output = r'C:\Users\Yinfinity\Desktop\nespresso_auto_prs\data\nespresso_rpt_ini_202401.pptx'


if __name__ == '__main__':

    run(prs_data_load, prs_load, prs_output)