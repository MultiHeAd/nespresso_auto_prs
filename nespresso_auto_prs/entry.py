# -*- coding:utf-8 -*-

from slides.generate_slides import run
import os

project_folder=os.path.dirname(os.path.realpath(__file__))
prs_data_load = os.path.join(project_folder,'./data/nespresso_rpt_data_202401.xlsx')
prs_load = os.path.join(project_folder,'./data/nespresso_rpt_demo.pptx')
prs_output = os.path.join(project_folder,'./data/nespresso_rpt_ini_202401.pptx')


if __name__ == '__main__':

    run(prs_data_load, prs_load, prs_output)