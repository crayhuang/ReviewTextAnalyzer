#!/usr/bin/python2.7
# -*- coding: utf-8 -*-

__author__ = 'Jeff'

import xlrd
import csv
import re
import time
import datetime
import sys
import os

reload(sys)
sys.setdefaultencoding("utf-8")

# Bad word files path
appliances_bad_words_path = './BadWordFiles/APPLIANCES.txt'
baby_care_bad_words_path = './BadWordFiles/BABY CARE.txt'
fabric_care_bad_words_path = './BadWordFiles/FABRIC CARE.txt'
feminine_care_bad_words_path = './BadWordFiles/FEMININE CARE.txt'
hair_care_bad_words_path = './BadWordFiles/HAIR CARE.txt'
oral_care_bad_words_path = './BadWordFiles/ORAL CARE.txt'
personal_care_bad_words_path = './BadWordFiles/PERSONAL CARE.txt'
prestige_bad_words_path = './BadWordFiles/PRESTIGE.txt'
shave_care_bad_words_path = './BadWordFiles/SHAVE CARE.txt'
skin_care_bad_words_path = './BadWordFiles/SKIN CARE.txt'
air_care_bad_words_path = './BadWordFiles/AIR CARE.txt'

# Keyword files path
appliances_keywords_path = "./KeywordFiles/APPLIANCES.xlsx"
baby_care_keywords_path = "./KeywordFiles/Baby Care.xlsx"
fabric_care_keywords_path = "./KeywordFiles/fabric care.xlsx"
feminine_care_keywords_path = "./KeywordFiles/Femine Care.xlsx"
hair_care_keywords_path = "./KeywordFiles/HAIR CARE.xlsx"
oral_care_keywords_path = "./KeywordFiles/ORAL CARE.xlsx"
personal_care_keywords_path = "./KeywordFiles/PERSONAL CARE.xlsx"
prestige_keywords_path = "./KeywordFiles/PRESTIGE.xlsx"
shave_care_keywords_path = "./KeywordFiles/SHAVE CARE.xlsx"
skin_care_keywords_path = "./KeywordFiles/SKIN CARE.xlsx"
air_care_keywords_path = "./KeywordFiles/AIR CARE.xlsx"

# Exceptional files path
appliances_exceptional_word_path = './ExceptionalWordFiles/APPLIANCES.txt'
baby_care_exceptional_word_path = './ExceptionalWordFiles/BABY CARE.txt'
fabric_care_exceptional_word_path = './ExceptionalWordFiles/FABRIC CARE.txt'
feminine_care_exceptional_word_path = './ExceptionalWordFiles/FEMININE CARE.txt'
hair_care_exceptional_word_path = './ExceptionalWordFiles/HAIR CARE.txt'
oral_care_exceptional_word_path = './ExceptionalWordFiles/ORAL CARE.txt'
personal_care_exceptional_word_path = './ExceptionalWordFiles/PERSONAL CARE.txt'
prestige_exceptional_word_path = './ExceptionalWordFiles/PRESTIGE.txt'
shave_care_exceptional_word_path = './ExceptionalWordFiles/SHAVE CARE.txt'
skin_care_exceptional_word_path = './ExceptionalWordFiles/SKIN CARE.txt'
air_care_exceptional_word_path = './ExceptionalWordFiles/AIR CARE.txt'

sentence_cutting_path = 'sentence_cutting.txt'
mean_conversions_path = 'mean_conversion.txt'


re_pattern = "\,|，|\.|。|\!|！|\?|？|\；|;|\=|：|\ |、|\~|\……|\--| "


class ReviewTextAnalyzer:
    def __init__(self):
        self.category = None
        self.input_file_path = ''
        self.output_file_path = './OutputFiles/'
        self.bad_word_path = ''
        self.keyword_path = ''
        self.exceptional_path = ''
        self.keywords_sheet = None
        self.bad_word_list = []
        self.exceptional_word_list = []
        self.mean_conversion_list = []
        self.sentence_cutting_list = []
        self.results = []
        self.header = []
        self.start_time = 0
        self.end_time = 0

    def init_category(self, category):
        self.category = category
        if 'APPLIANCES' == self.category:
            self.init_appliances()
        if 'BABY_CARE' == self.category:
            self.init_baby_care()
        if 'FABRIC_CARE' == self.category:
            self.init_fabric_care()
        if 'FEMININE_CARE' == self.category:
            self.init_feminine_care()
        if 'HAIR_CARE' == self.category:
            self.init_hair_care()
        if 'ORAL_CARE' == self.category:
            self.init_oral_care()
        if 'PERSONAL_CARE' == self.category:
            self.init_personal_care()
        if 'PRESTIGE' == self.category:
            self.init_prestige()
        if 'SHAVE_CARE' == self.category:
            self.init_shave_care()
        if 'SKIN_CARE' == self.category:
            self.init_skin_care()
        if 'AIR_CARE' == self.category:
            self.init_air_care()

    def set_input_file_path(self, input_file_path):
        self.input_file_path = input_file_path

    def load_keywords_sheet(self):
        wb = xlrd.open_workbook(self.keyword_path)
        self.keywords_sheet = wb.sheet_by_index(0)

    def load_bad_words(self):
        bad_words_reader = open(self.bad_word_path, 'r')
        bad_words = bad_words_reader.xreadlines()
        for bw in bad_words:
            self.bad_word_list.append(bw.strip())

    def load_mean_conversions(self):
        mean_conversions_reader = open(mean_conversions_path, 'r')
        mean_conversions = mean_conversions_reader.xreadlines()
        for mc in mean_conversions:
            self.mean_conversion_list.append(mc.strip())

    def load_sentence_cuttings(self):
        sentence_cuttings_reader = open(sentence_cutting_path, 'r')
        sentence_cuttings = sentence_cuttings_reader.xreadlines()
        for sc in sentence_cuttings:
            self.sentence_cutting_list.append(sc.strip())

    def load_exceptional_words(self):
        exceptional_words_reader = open(self.exceptional_path)
        exceptional_words = exceptional_words_reader.xreadlines()
        for ew in exceptional_words:
            self.exceptional_word_list.append(ew.strip())

    def analyze(self):
        self.start_time = datetime.datetime.now()
        self.load_keywords_sheet()
        self.load_bad_words()
        self.load_mean_conversions()
        self.load_sentence_cuttings()
        self.load_exceptional_words()

        print self.input_file_path
        i = 0
        file_path_without_ext = os.path.splitext(self.input_file_path)[0]
        file_path_parts = os.path.split(file_path_without_ext)
        self.output_file_path += file_path_parts[-1]
        print self.output_file_path
        with open(self.input_file_path, 'rU') as input_file:
            input_reader = csv.reader((line.replace('\0', '') for line in input_file), dialect=csv.excel_tab,
                                      delimiter=',')
            input_col_len = len(input_reader.next())
            for row in input_reader:
                review_text = row[12]
                review_text = review_text.strip()
                handling_text = review_text
                for ew in self.exceptional_word_list:
                    ew = ew.strip().decode('GBK').encode('utf-8')
                    handling_text = handling_text.replace(ew, '')

                for cuttings in self.sentence_cutting_list:
                    if cuttings.strip() in handling_text:
                        handling_text = handling_text.replace(cuttings, '*')

                split_texts = re.split(re_pattern, handling_text)
                has_kw = False
                existing_review_subtype_array = []
                for sub_text in split_texts:
                    keywords_len = self.keywords_sheet.nrows - 1
                    current_row = -1
                    for kw in self.keywords_sheet.col_values(4):
                        current_row += 1
                        review_type = self.keywords_sheet.cell_value(current_row, 2).encode('utf-8')
                        review_subtype = self.keywords_sheet.cell_value(current_row, 3).encode('utf-8')
                        if review_subtype in existing_review_subtype_array:
                            continue
                        if isinstance(kw, float):
                            kw = str(kw)
                        kw = kw.encode('utf-8')
                        if kw in sub_text:
                            score = '5'
                            bw = ''
                            for bad_word in self.bad_word_list:
                                # Remove space and \n from bad word
                                if len(bad_word) <= 0:
                                    continue
                                bad_word = bad_word.strip().decode('GBK').encode('utf-8')
                                if bad_word in sub_text:
                                    bw = bad_word
                                    score = '1'
                                    for mc in self.mean_conversion_list:
                                        mc = mc.strip()
                                        if len(mc) > 0:
                                            if mc in sub_text.replace(bad_word, '') and sub_text.find(mc) < sub_text.find(bad_word):
                                                score = '5'
                                                break
                            clone_row = row[:]
                            clone_row.append(review_type)
                            clone_row.append(review_subtype)
                            existing_review_subtype_array.append(review_subtype)
                            # clone_row.append(sub_text)
                            # # clone_row.append(kw)
                            # clone_row.append(bw)
                            clone_row[5] = kw
                            clone_row[11] = score
                            has_kw = True
                            self.results.append(clone_row)
                if not has_kw:
                    clone_row = row[:]
                    clone_row.append('others')
                    clone_row.append('others')
                    self.results.append(clone_row)

    def output(self):
        # self.output_file_path += self.category
        self.output_file_path += '_'
        self.output_file_path += str(time.strftime('%Y%m%d-%H%M%S'))
        self.output_file_path += '.csv'
        self.header = ['brand', 'category', 'is_competitor', 'manufacturer', 'market', 'matched_keywords', 'online_store',
                       'product_description', 'report_date', 'retailer_product_code', 'review_date', 'review_rating',
                       'review_text', 'review_title', 'sub_category', 'time_of_publication', 'upc', 'url', 'unique_ID',
                       'review type', 'review subtype']
        with open(self.output_file_path, 'wb') as output_file:
            output_writer = csv.writer(output_file, delimiter=",", quotechar='"', quoting=csv.QUOTE_ALL)
            output_writer.writerow(self.header)
            for rs in self.results:
                output_writer.writerow(rs)

        self.end_time = datetime.datetime.now()
        print "total cost time: %s second" % (str((self.end_time - self.start_time).seconds))
        print "Finish!!!"

    def init_personal_care(self):
        self.keyword_path = personal_care_keywords_path
        self.bad_word_path = personal_care_bad_words_path
        self.exceptional_path = personal_care_exceptional_word_path

    def init_appliances(self):
        self.keyword_path = appliances_keywords_path
        self.bad_word_path = appliances_bad_words_path
        self.exceptional_path = appliances_exceptional_word_path

    def init_baby_care(self):
        self.keyword_path = baby_care_keywords_path
        self.bad_word_path = baby_care_bad_words_path
        self.exceptional_path = baby_care_exceptional_word_path

    def init_fabric_care(self):
        self.keyword_path = fabric_care_keywords_path
        self.bad_word_path = fabric_care_bad_words_path
        self.exceptional_path = fabric_care_exceptional_word_path

    def init_feminine_care(self):
        self.keyword_path = feminine_care_keywords_path
        self.bad_word_path = feminine_care_bad_words_path
        self.exceptional_path = feminine_care_exceptional_word_path

    def init_hair_care(self):
        self.keyword_path = hair_care_keywords_path
        self.bad_word_path = hair_care_bad_words_path
        self.exceptional_path = hair_care_exceptional_word_path

    def init_oral_care(self):
        self.keyword_path = oral_care_keywords_path
        self.bad_word_path = oral_care_bad_words_path
        self.exceptional_path = oral_care_exceptional_word_path

    def init_prestige(self):
        self.keyword_path = prestige_keywords_path
        self.bad_word_path = prestige_bad_words_path
        self.exceptional_path = prestige_exceptional_word_path

    def init_shave_care(self):
        self.keyword_path = shave_care_keywords_path
        self.bad_word_path = shave_care_bad_words_path
        self.exceptional_path = shave_care_exceptional_word_path

    def init_skin_care(self):
        self.keyword_path = skin_care_keywords_path
        self.bad_word_path = skin_care_bad_words_path
        self.exceptional_path = skin_care_exceptional_word_path

    def init_air_care(self):
        self.keyword_path = air_care_keywords_path
        self.bad_word_path = air_care_bad_words_path
        self.exceptional_path = air_care_exceptional_word_path
