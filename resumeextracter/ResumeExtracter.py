# 读取、操作简历的类
# import module
import os
import re
from docx import Document
from openpyxl import Workbook


class ResumeReader:
    # 简历阅读器类，可迭代，向外提供简历的诸多要素
    # 实例方法
    def __init__(self, file_path):
        # 实例属性
        # self.person_name =
        # self.corps = [] # 公司
        # self.job_titles = [] # 职位
        # self.exp = [] # 经历
        self.file_path = file_path  # 目标文件夹
        self.__document = Document(file_path)
        self.work_exp_text = self.__get_experience_parts()
        self.__company_names_index = self.__get_company_names_index()

    # extract person name from file name
    def get_person_name(self):
        name_reg = re.compile('[\u4e00-\u9fa5]*简历')
        name_raw = name_reg.findall(self.file_path)
        stopwords_reg = re.compile('个人|中文|英文|简历')
        person_name = re.sub(stopwords_reg, '', name_raw[0])
        return person_name

    # locate the experience part of the resume
    def __get_experience_parts(self):
        start_index = -1
        end_index = -1
        flag_reg = re.compile("(?:工作|实习|项目)(?:经验|经历)")
        # content of the resume
        doc_paras = [para.text for para in self.__document.paragraphs]
        # iterate the content
        for para in doc_paras:
            # find the start position only if it is not found
            if start_index == -1:
                if re.match(flag_reg, para) is not None:
                    start_index = doc_paras.index(para) + 1
                elif '公司' in para:
                    start_index = doc_paras.index(para)
                    print('find comp' + para)
                else:
                    continue
            # then find the end position
            if len(para) <= 6:
                end_index = doc_paras.index(para)
            else:
                continue
        # return texts about working experience
        if end_index != -1:
            return doc_paras[start_index:end_index]
        else:
            return doc_paras[start_index:]

    # extract company name index from resume text
    def __get_company_names_index(self):
        comp_index_list = []
        # only iterate working experience part
        exp_text_list = self.work_exp_text
        # regex
        comp_reg = re.compile('[\S]*(?:公司|银行)\s')
        # iterate
        for i in range(len(exp_text_list)):
            exp_text = exp_text_list[i]
            if re.match(comp_reg, exp_text) is not None:
                comp_index_list.append(i)

        return comp_index_list

    # extract company name from resume text according to the index
    def get_company_names(self):
        return [self.work_exp_text[i] for i in self.__company_names_index]

    def get_work_exps(self):
        single_work_exp = []
        work_exps = []
        for i in range(len(self.__company_names_index)):
            index = self.__company_names_index[i] + 1
            # not the last
            if i < len(self.__company_names_index) - 1:
                while index < self.__company_names_index[i + 1]:
                    single_work_exp.append(self.work_exp_text[index])
                    index += 1
                work_exps.append(''.join(single_work_exp))
                single_work_exp = []
            else:
                while index < len(self.work_exp_text):
                    single_work_exp.append(self.work_exp_text[index])
                    index += 1
                work_exps.append(''.join(single_work_exp))

        return work_exps


class ResumeExtracter:
    # 简历提取类，向外提供一个将目标文件夹内的所有docx简历文件导出成excel的方法
    # 实例方法
    def __init__(self, dir_path):
        self.dir_path = dir_path  # 简历所在文件夹

    # 私有方法
    # 判断目标文件夹是否存在
    def __is_dir_exists(self):
        return os.path.exists(self.dir_path)

    def __get_resume_names(self):
        if self.__is_dir_exists():  # 如果目标文件夹存在
            resume_names = list(filter(lambda f: str(f).endswith('docx'), os.listdir(self.dir_path)))  # 只保留docx文件
            return resume_names
        else:
            raise Exception('Appointed Directory is not existed!')

    def to_excel(self, file_name, to_path):
        # create a new excel workbook
        wb = Workbook()
        # crate a new excel sheet
        ws = wb.active
        # create title
        ws_title = ['姓名', '工作单位', '工作经历']
        ws.append(ws_title)
        # 循环
        for resume_name in self.__get_resume_names():
            resume_path = self.dir_path + resume_name if self.dir_path.endswith(
                '/') else self.dir_path + '/' + resume_name
            # 读取简历文件
            resume_reader = ResumeReader(resume_path)
            # get person name
            person_name = resume_reader.get_person_name()
            # company name
            company_name_list = resume_reader.get_company_names()
            # work experience
            work_exps_list = resume_reader.get_work_exps()
            for i in range(len(company_name_list)):
                output_context = [person_name, company_name_list[i], work_exps_list[i]]
                # writ to excel file
                ws.append(output_context)
        # save file
        file_path = to_path + file_name if to_path.endswith('/') else to_path + '/' + file_name
        # check existence
        incre_suffix = 0
        file_name_wo_suffix = re.sub(r'\.xlsx', '', file_name)
        while os.path.isfile(file_path):
            incre_suffix += 1
            file_path = to_path + file_name_wo_suffix + '-' + str(incre_suffix) + '.xlsx' \
                if to_path.endswith('/') else to_path + '/' + file_name_wo_suffix + '-' + str(incre_suffix) + '.xlsx'
        # save
        wb.save(file_path)
        # close
        wb.close()
