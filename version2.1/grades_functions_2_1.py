#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Created on Sun Jun 27 15:43:22 2021

@author: huangmu
"""
import sys
import pandas as pd

def get_data(file, semesters, term_n):
    global df , terms
    terms = semesters
    df=[]
    for i in range(term_n):
        df.append(pd.read_excel(file,semesters[i]))
    
def data_term(term_n, option):
    global credits_, grades, points, courses, courses_opt
    
    credits_ = []
    grades = []
    points = []
    courses = []
    courses_opt = []
    option_courses(option)
    
    for j in range(term_n):
        credit = df[j]['学分']
        grade = df[j]['分数']
        point = trans(grade)
        sort = df[j]['课程类型']
        name = df[j]['课程']
        
        grade_avg = [0, 0]
        point_avg = [0, 0]
        course_count = [0, 0]
        credit_count = [0, 0]
        credit_count[0] = credit.sum()
        course_count[0] = sort.count()
        
        for i in range(len(sort)):
            grade_avg[0] += grade[i] * credit[i]
            point_avg[0] += point[i] * credit[i]
            if '选修' not in sort[i]:
                course_count[1] += 1
                credit_count[1] += credit[i]
                grade_avg[1] += grade[i] * credit[i]
                point_avg[1] += point[i] * credit[i]
            if '选修' in sort[i]:
                courses_opt.append([name[i], trans_sort(sort[i]), credit[i]])
            for k in range(len(data_opt)):
                if data_opt[k][0] in sort[i]:
                    data_opt[k][1] += credit[i]
                    data_opt[k][2] += 1
                
            
        credits_.append(credit_count)
        courses.append(course_count)
        grades.append([grade_avg[0]/credit_count[0], grade_avg[1]/credit_count[1]])
        points.append([point_avg[0]/credit_count[0], point_avg[1]/credit_count[1]])
        
    #处理选修课
    courses_opt = pd.DataFrame(courses_opt)
    courses_opt = courses_opt.rename(columns = {0:'name', 1:'sort', 2:'credit'})

def option_courses(option):
    global data_opt
    n = len(option)
    data_opt = []
    
    for i in range(n):
        data_opt.append([option[i][0], 0, 0, option[i][1], option[i][2]])
    

def show_outcomes(term_n, terms):
    for i in range(term_n):
        print('-'*50)
        print(terms[i])
        print('-'*20)
        print('本学期所修学分:  {:,.2f}(必修) , {:,.2f}(总共)'.format(credits_[i][1],credits_[i][0]))
        print('本学期所修课程:  {:,.0f}门(必修) , {:,.0f}门(总共)'.format(courses[i][1],courses[i][0]))
        print('必修课程:   grades: {:,.3f} , point: {:,.3f}'.format(grades[i][1] ,points[i][1]))
        print('所有课程:   grades: {:,.3f} , point: {:,.3f}'.format(grades[i][0] ,points[i][0]))
        print('')
    print('-'*50)
    print('直至' + terms[term_n-1])
    print('共修学分/分:     {:,.2f}(必修) , {:,.2f}(总共)'.format(get_sum(credits_, 1),get_sum(credits_, 0)))
    print('共修课程/门:        {:,.0f}(必修) ,     {:,.0f}(总共)'.format(get_sum(courses, 1),get_sum(courses, 0)))
    print('必修课程:  grades: {:,.3f} , point: {:,.3f}'.format(get_avg(grades, 1) ,get_avg(points, 1)))
    print('所有课程:  grades: {:,.3f} , point: {:,.3f}'.format(get_avg(grades, 0) ,get_avg(points, 0)))
    print('-'*10,'选修课情况','-'*10)
    for i in range(len(data_opt)):
        print(data_opt[i][0]+'课: 已修', data_opt[i][2],'门  ',data_opt[i][1],'学分 ',end='     ')
        credit = data_opt[i][3] - data_opt[i][1]
        course = data_opt[i][4] - data_opt[i][2]
        if credit <= 0 and course <= 0:
            print('(已达标)')
        else:
            print('还差  ', end = '')
            if credit > 0:
                print(credit, '学分', end = '  ')
            if course > 0:
                print(course, '门')
            else:
                print('')
    print('-'*50)
    
def get_avg(data, k):
    n = len(data)
    avg, credit_ = 0, 0
    for i in range(n):
        avg += data[i][k] * credits_[i][k]
        credit_ += credits_[i][k]
    return avg/credit_

def get_sum(data, k):
    n = len(data)
    ret = 0
    for i in range(n):
        ret += data[i][k]
    return ret

def trans_sort(text):
    if '通识' in text:
        return '通识'
    elif '实践' in text:
        return '实践'
    else:
        return '专业'

def trans(grades):
    table_grades = [60,64, 68, 72, 75, 78, 82, 85, 90]
    table_point = [1.3, 1.5, 2, 2.3, 2.7, 3, 3.3, 3.7, 4]
    n = len(grades)
    ret = []
    for j in range(n):
        m = grades[j]
        n = 0
        for i in range(len(table_grades)):
            if(m >= table_grades[i]):
                n = table_point[i]

        ret.append( n)
    return ret
        
