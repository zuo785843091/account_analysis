'''
Created on 2019年6月13日

@author: Z
'''
import xlwings as xw
from configs import configs
from logging_config import logging
import xw_funs
import time

if __name__ == '__main__':
    excel_configs = configs()
    
    wb1 = xw.Book(excel_configs.account_analysis_book)
    wb2 = xw.Book(excel_configs.non_verification_book)
    wb3 = xw.Book(excel_configs.payable_book)
    #wb1.app.screen_updating=False
    #wb2.app.screen_updating=False
    #wb3.app.screen_updating=False
    try:
        # 检查 ‘未核销’
        logging.info('=====================开始更新“未核销”=======================')
        subjects, subjects_tables = xw_funs.get_subjects(wb1, sht_name = excel_configs.template_sheet, item_analysis = '未核销')
        
        for subject, subjects_table in zip(subjects, subjects_tables):
            subject = subject.strip()
            subjects_table = subjects_table.strip()
            dict_suppliers = xw_funs.check_non_verification_table(wb2, subject, sht_name = excel_configs.non_verification_sheet)
            if len(dict_suppliers) == 0:
                logging.warning("未在'未核销表'中查找到'{subject}'科目,请手动确认!!!".format(subject = subject))
                continue
            t_start = time.time()
            logging.info("‘{subjects_tb}’，共{suppliers}条正在更新...".format(subjects_tb = subjects_table, suppliers = len(dict_suppliers)))
            xw_funs.update_account_analysis_table(wb1, subjects_table, dict_suppliers)
            t_end = time.time()
            logging.info("‘{subjects_tb}’更新完毕！ 耗时：{time_used:.2f}秒".format(subjects_tb = subjects_table, time_used = t_end - t_start))
        
        # 检查 ‘应付’
        logging.info('======================开始更新“应付”========================')
        subjects, subjects_tables = xw_funs.get_subjects(wb1, sht_name = excel_configs.template_sheet, item_analysis = '应付')
    
        for subject, subjects_table in zip(subjects, subjects_tables):
            subject = subject.strip()
            subjects_table = subjects_table.strip()
            dict_suppliers = xw_funs.check_payable_table(wb3, subject, sht_name = excel_configs.payable_sheet)
            if len(dict_suppliers) == 0:
                logging.warning("未在'应付表'中查找到'{subject}'科目,请手动确认!!!".format(subject = subject))
                continue
            t_start = time.time()
            logging.info("‘{subjects_tb}’，共{suppliers}条正在更新...".format(subjects_tb = subjects_table, suppliers = len(dict_suppliers)))
            xw_funs.update_account_analysis_table(wb1, subjects_table, dict_suppliers) 
            t_end = time.time()
            logging.info("‘{subjects_tb}’更新成功！ 耗时：{time_used:.2f}秒".format(subjects_tb = subjects_table, time_used = t_end - t_start))
            
        logging.info('=======================更新完毕!=======================')
    finally:
        wb1.sheets[excel_configs.template_sheet].activate()