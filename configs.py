'''
Created on 2019年6月12日

@author: Zcl
'''





class configs(object):
    '''
    classdocs
    '''


    def __init__(self):
        '''
        Constructor
        '''
        
        self.account_analysis_book = r'往来分析表-201904（胡婷彬）.xls'
        self.non_verification_book = r'201905-CUC_未核销预付款明细表_030619.xls'
        self.payable_book = r'201905-CUC_应付帐款和其他应付帐款报表_030619.xls'
        
        self.template_sheet = '模板'
        self.non_verification_sheet = '201905-CUC_未核销预付款明细表_030619'
        self.payable_sheet = '201905-CUC_应付帐款和其他应付帐款报表_030619'
        self.MAX_ROW_NUM = 5000
        self.MAX_COL_NUM = 300