#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Created on Tue Jun 22 14:01:37 2021

用于计算成绩
@author: huangmu
"""

from grades_functions import *

if __name__ == "__main__":
    #文件与分表
    file = '个人成绩记录.xls'
    semesters = ['大一上学期','大一下学期','大二上学期','大二下学期','大三上学期','大三下学期','大四上学期','大四下学期']
    
    #选择是否写入文件,1为是
    out_to_file = 1
    if out_to_file == 1:
        file_out = '成绩分析.txt'
        sys.stdout = open(file_out, mode = 'w')
    
    #选择读入学期数
    term_n = 4
    #读入数据
    get_data(file, semesters, term_n)
    
    #成绩分析
    data_term(term_n)
    
    #结果展示
    show_outcomes(term_n, semesters)
        
        
    
