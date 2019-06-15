'''
Created on 2019年6月12日

@author: Zcl
'''
import xlwings as xw
import numpy as np
from configs import configs
from logging_config import logging


MAX_ROW_NUM = 5000
MAX_COL_NUM = 300

def open_book(excel_file):
    wb = xw.Book(excel_file)
    return wb

# 将字母列转成数字 0代表第A列
def col_to_num(col):
    dict_col_to_num = {'A':1, 'B':2, 'C':3, 'D':4, 'E':5, 'F':6, 'G':7, 'H':8, 'I':9, 'J':10, \
                       'K':11, 'L':12, 'M':13, 'N':14, 'O':15, 'P':16, 'Q':17, 'R':18, 'S':19, 'T':20, \
                       'U':21, 'V':22, 'W':23, 'X':24, 'Y':25, 'Z':26}
    num = 0    
    index_s = 0
    for s in col[-1: :-1]:
        num += (26 ** index_s) * dict_col_to_num[s]
        index_s += 1    
   
    return num - 1

# 将数字转成字母列 0代表第A列
def num_to_col(num):
    dict_num_to_col = {1:'A', 2:'B', 3:'C', 4:'D', 5:'E', 6:'F', 7:'G', 8:'H', 9:'I', 10:'J', \
                       11:'K', 12:'L', 13:'M', 14:'N', 15:'O', 16:'P', 17:'Q', 18:'R', 19:'S', 20:'T', \
                       21:'U', 22:'V', 23:'W', 24:'X', 25:'Y', 26:'Z', 0:''}
    num += 1
    (multiple, residue) = divmod(num, 26)
    col = dict_num_to_col[multiple] + dict_num_to_col[residue]
    
    return col

def cal_stage(aging_days):    
    aging_mouths = aging_days // 30
    
    if aging_mouths >= 0 and aging_mouths <= 6:
        stage = 0
    elif aging_mouths >= 7 and aging_mouths <= 12:
        stage = 1
    elif aging_mouths >= 12 and aging_mouths <= 24:  #1-2年
        stage = 2
    elif aging_mouths >= 25 and aging_mouths <= 36:  #2-3年
        stage = 3
    elif aging_mouths >= 37 and aging_mouths <= 48:  #3-4年
        stage = 4  
    elif aging_mouths >= 49 and aging_mouths <= 60:  #4-5年
        stage = 5
    elif aging_mouths >= 61:  #5年以上
        stage = 6
    else:
        raise RuntimeError("帐龄不在范围内! {aging_days}".format(aging_days = aging_days))
        
    return stage

# 检查是否满足最大最小行
def check_max_rc(arr, sht_name):
    if len(arr.shape) < 2:
        raise RuntimeError("'{sht_name}表'的行列数超出最小规定范围, 请检查表格。\n \
            shape = {shape}".format(sht_name=sht_name, shape=str(arr.shape)))
    
    if arr.shape[0] > MAX_ROW_NUM or arr.shape[1] > MAX_COL_NUM:
        raise RuntimeError("'{sht_name}表'的行列数超出最大规定范围，请删除多余的空行或修改范围值。\n \
            shape = {shape}".format(sht_name=sht_name, shape=str(arr.shape)))
        
    elif arr.shape[0] < 2 or arr.shape[1] < 2:
        raise RuntimeError("'{sht_name}表'的行列数超出最小规定范围, 请检查表格。\n \
            shape = {shape}".format(sht_name=sht_name, shape=str(arr.shape)))

def delete_npstr_space(arr):
    for i in range(len(arr)):        
        arr[i] = str(arr[i]).strip()

    return arr


# 获取模板表中的对应item_analysis（‘未核销’或‘应付’）的条目
def get_subjects(wb, sht_name = '模板', item_analysis = '未核销', sht_range = None):
    col_items = '条目'
    col_subjects = '科目'
    col_subjects_table = '科目表名'
    
    sht = wb.sheets[sht_name]
    if sht_range != None:
        arr = sht.range(sht_range).options(np.array).value
    else:
        arr = sht.used_range.options(np.array).value
        check_max_rc(arr, sht_name)

    # 获取列索引
    items_col = np.argwhere(arr[0] == col_items)[0, 0]
    subjects_col = np.argwhere(arr[0] == col_subjects)[0, 0]
    subjects_table_col = np.argwhere(arr[0] == col_subjects_table)[0, 0]
    
    # 获取条目列
    arr_items = arr[:, items_col]
    
    # 获取item_analysis对应的行索引
    analysis_row = np.argwhere(arr_items == item_analysis)[:, 0]
    
    subjects = arr[analysis_row, subjects_col]
    subjects_tables = arr[analysis_row, subjects_table_col]
    
    if len(np.argwhere(subjects == 'nan')) >= 1:
        index_nan = np.argwhere(subjects == 'nan')
        subjects = np.delete(subjects, index_nan)
        subjects_tables = np.delete(subjects_tables, index_nan)
        logging.warning("'{sht_name}表'中科目列{row}行存在空单元格，请确认！".format(sht_name=sht_name, row = index_nan))
    elif len(np.argwhere(subjects_tables == 'nan')) >= 1:
        index_nan = np.argwhere(subjects_tables == 'nan')
        subjects = np.delete(subjects, index_nan)
        subjects_tables = np.delete(subjects_tables, index_nan)
        logging.warning("'{sht_name}表'中科目表名列{row}行存在空单元格，请确认！".format(sht_name=sht_name, row = index_nan))

    # dict_subjects = dict.fromkeys(subjects, subjects_tables)
    
    return subjects, subjects_tables
    
def get_supplier_data(wb, subject, sht_name, col_subjects_name, col_supplier_name, \
                      col_non_verification_amount, col_aging_days, col_remark, \
                      sht_range = None):
    
    sht = wb.sheets[sht_name]
    if sht_range != None:
        arr = sht.range(sht_range).options(np.array).value
    else:
        arr = sht.used_range.options(np.array).value
        check_max_rc(arr, sht_name)
        
    subjects_name_col = np.argwhere(arr == col_subjects_name)[0, 1]
    supplier_name_col = np.argwhere(arr == col_supplier_name)[0, 1]
    non_verification_amount_col = np.argwhere(arr == col_non_verification_amount)[0, 1]
    aging_days_col = np.argwhere(arr == col_aging_days)[0, 1]
    remark_col = np.argwhere(arr == col_remark)[0, 1]
    
    arr_subjects = arr[:, subjects_name_col]    
    arr_subjects = delete_npstr_space(arr_subjects)
    subject_rows = np.argwhere(arr_subjects == subject)[:, 0]
    
    dict_suppliers = {}
    for i in range(len(subject_rows)):
        supplier = arr[subject_rows[i], supplier_name_col].strip()
        non_verification_amount = arr[subject_rows[i], non_verification_amount_col]
        aging_days = arr[subject_rows[i], aging_days_col]
        remark = arr[subject_rows[i], remark_col]
        
        if supplier not in dict_suppliers.keys():
            dict_suppliers[supplier] = [[non_verification_amount, aging_days, remark]]
        else:
            values = dict_suppliers.get(supplier)
            values.append([non_verification_amount, aging_days, remark])
            dict_suppliers[supplier] = values            

    return dict_suppliers
    
# 从未核销表提取对应科目的供应商数据
def check_non_verification_table(wb, subject, sht_name = '未核销表', sht_range = None):
    col_subjects_name = '科目描述'
    col_supplier_name = '供应商名称'
    col_non_verification_amount = '未核销本位币金额_入帐'
    col_aging_days = '帐龄天数'
    col_remark = '款项性质'
    
    dict_suppliers = get_supplier_data(wb, subject, sht_name, col_subjects_name, col_supplier_name, \
                                       col_non_verification_amount, col_aging_days, col_remark, \
                                       sht_range)

    return dict_suppliers

# 从应付表提取对应科目的供应商数据
def check_payable_table(wb, subject, sht_name = '应付表', sht_range = None):
    col_subjects_name = '科目描述'
    col_supplier_name = '供应商'
    col_pay_amount = '总计'
    col_aging_days = '帐龄'
    col_remark = '发票摘要'
    
    dict_suppliers = get_supplier_data(wb, subject, sht_name, col_subjects_name, col_supplier_name, \
                                       col_pay_amount, col_aging_days, col_remark, \
                                       sht_range)
    
    return dict_suppliers

def update_account_analysis_table(wb, sht_name, dict_suppliers, sht_range = None):    
    col_supplier_name = '非关联公司名称'
    col_stage_1 = '0-6个月'
    row_last = '合计'
    
    sht = wb.sheets[sht_name]
    if sht_range != None:
        arr = sht.range(sht_range).options(np.array).value
    else:
        arr = sht.used_range.options(np.array).value
        check_max_rc(arr, sht_name)
    supplier_name_col = np.argwhere(arr == col_supplier_name)[0, 1]
    supplier_name_row = np.argwhere(arr == col_supplier_name)[0, 0]    
    stage_1_col = np.argwhere(arr == col_stage_1)[0, 1]
    last_row = int(np.argwhere(arr == row_last)[0, 0])
    
    # 清空遗留数据 J* : P*
    fsc = num_to_col(stage_1_col)
    lsc = num_to_col(stage_1_col + 6)
    clear_range = "{fsc}{r_start}:{lsc}{r_end}".format(fsc = fsc, lsc = lsc, r_start = supplier_name_row+2, r_end = last_row)
    #clear_range = "J{r_start}:P{r_end}".format(r_start = supplier_name_row+2, r_end = last_row)
    sht.api.Range(clear_range).ClearContents()
    
    # 更新已有的供应商
    for i in range(supplier_name_row + 1, last_row):
        supplier = str(arr[i, supplier_name_col]).strip()
        values = dict_suppliers.get(supplier)
        
        if values:
            non_verification_amount = []
            stage = []
            for v in values:
                aging_days = v[1]
                non_verification_amount.append(v[0])
                stage.append(cal_stage(aging_days))            
            
            for s in set(stage):
                nva = 0
                for j in range(len(stage)):
                    if s == stage[j]:
                        nva += non_verification_amount[j]
                sht[i, int(stage_1_col + s)].value = nva
            
            sht[i, int(stage_1_col - 1)].value = sum(non_verification_amount)
        
            del dict_suppliers[supplier]
            
    # 新增供应商
    if len(dict_suppliers) > 0:
        for supplier in dict_suppliers: 
            values = dict_suppliers.get(supplier)
            non_verification_amount = []
            stage = []            
                       
            insert_row = last_row
            sht.api.Rows(insert_row + 1).Insert()
            sht[insert_row, int(stage_1_col)].value = np.array([None] * 7)
            
            sht[insert_row, 0].value = arr[supplier_name_row + 1, 0]
            sht[insert_row, 1].value = arr[supplier_name_row + 1, 1]
            sht[insert_row, 5].value = arr[supplier_name_row + 1, 5]            
            sht[insert_row, int(supplier_name_col)].value = supplier
            sht[insert_row, int(supplier_name_col + 1)].value = values[0][2]
            

            for v in values:
                aging_days = v[1]
                non_verification_amount.append(v[0])
                stage.append(cal_stage(aging_days))
            
            for s in set(stage):
                nva = 0
                for j in range(len(stage)):
                    if s == stage[j]:
                        nva += non_verification_amount[j]
                sht[insert_row, int(stage_1_col + s)].value = nva
            sht[insert_row, int(stage_1_col - 1)].value = sum(non_verification_amount)
        
            last_row += 1
            logging.info('    新增：' + supplier)

def copy_last_account(wb, sht_name, sht_range = None):    
    col_supplier_name = '非关联公司名称'
    col_final_stage = '期末'
    row_last = '合计'
    

    sht = wb.sheets[sht_name]
    if sht_range != None:
        arr = sht.range(sht_range).options(np.array).value
    else:
        arr = sht.used_range.options(np.array).value
        check_max_rc(arr, sht_name)
    #supplier_name_col = np.argwhere(arr == col_supplier_name)[0, 1]
    supplier_name_row = np.argwhere(arr == col_supplier_name)[0, 0]
    final_stage_col = np.argwhere(arr == col_final_stage)[0, 1]    
    last_row = int(np.argwhere(arr == row_last)[0, 0])
    
    # 复制I列数据到H列
    fsc = num_to_col(final_stage_col)
    lsc = num_to_col(final_stage_col - 1)
    copy_form_range = "{fsc}{r_start}:{fsc}{r_end}".format(fsc = fsc, r_start = supplier_name_row+2, r_end = last_row)
    copy_to_range = "{lsc}{r_start}:{lsc}{r_end}".format(lsc = lsc, r_start = supplier_name_row+2, r_end = last_row)
    #copy_form_range = "I{r_start}:I{r_end}".format(r_start = supplier_name_row+2, r_end = last_row)
    #copy_to_range = "H{r_start}:H{r_end}".format(r_start = supplier_name_row+2, r_end = last_row)
    #sht.api.Range(copy_form_range).Copy(sht.api.Range(copy_to_range))  # 带格式复制
    # api方式速度更快
    sht.api.Range(copy_to_range).value = sht.api.Range(copy_form_range).value
    
    #for i in range(supplier_name_row + 1, last_row):
        #sht[i, int(final_stage_col - 1)].value = sht[i, int(final_stage_col)].value 
    #    sht[i, int(final_stage_col - 1)].value = arr[i, int(final_stage_col)]
    
    return True




if __name__ == '__main__':
    excel_configs = configs()    
    
    col = num_to_col(1000)
    print(col)
    
    num = col_to_num('A')
    print(num)
    print('end!')

        



