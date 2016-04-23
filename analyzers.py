#!/usr/bin/python2.7
# -*- coding: utf-8 -*-

__author__ = 'Jeff'

import xlrd
import csv
import re
import time

appliances_bad_words_path = './BadWordFiles/APPLIANCES.txt'
baby_care_bad_words_path = './BadWordFiles/APPLIANCES.txt'
fabric_care_bad_words_path = './BadWordFiles/FABRIC CARE.txt'
feminine_care_bad_words_path = './BadWordFiles/FEMININE CARE.txt'
hair_care_bad_words_path = './BadWordFiles/HAIR CARE.txt'
oral_care_bad_words_path = './BadWordFiles/ORAL CARE.txt'
personal_care_bad_words_path = './BadWordFiles/PERSONAL CARE.txt'
prestige_bad_words_path = './BadWordFiles/PRESTIGE.txt'
shave_care_bad_words_path = './BadWordFiles/SHAVE CARE.txt'
skin_care_bad_words_path = './BadWordFiles/SKIN CARE.txt'


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


sentence_cutting_path = 'sentence_cutting.txt'
mean_conversions_path = 'mean_conversion.txt'


re_pattern = "\,|，|\.|。|\!|！|\?|？|\；|;|\=|：|\ |、| "


class ReviewTextAnalyzer:
    def __init__(self):
        self.category = None
        self.input_file_path = ''
        self.output_file_path = './OutputFiles/'
        self.bad_word_path = ''
        self.keyword_path = ''
        self.keywords_sheet = None
        self.bad_word_list = []
        self.mean_conversion_list = []
        self.sentence_cutting_list = []
        self.results = []
        self.header = []

    def init_category(self, category):
        self.category = category
        if 'PERSONAL_CARE' == self.category:
            self.init_personal_care()
        if 'APPLIANCES' == self.category:
            self.init_appliances()
        if 'FABRIC_CARE' == self.category:
            self.init_fabric_care()
        if 'FEMININE_CARE' == self.category:
            self.init_feminine_care()
        if 'HAIR_CARE' == self.category:
            self.init_hair_care()
        if 'ORAL_CARE' == self.category:
            self.init_oral_care()
        if 'PRESTIGE' == self.category:
            self.init_prestige()
        if 'SHAVE_CARE' == self.category:
            self.init_shave_care()
        if 'SKIN_CARE' == self.category:
            self.init_skin_care()

    def set_input_file_path(self, input_file_path):
        self.input_file_path = input_file_path

    def load_keywords_sheet(self):
        wb = xlrd.open_workbook(self.keyword_path)
        self.keywords_sheet = wb.sheet_by_index(0)

    def load_bad_words(self):
        bad_words_reader = open(self.bad_word_path, 'r')
        bad_words = bad_words_reader.xreadlines()
        for bw in bad_words:
            self.bad_word_list.append(bw)

    def load_mean_conversions(self):
        mean_conversions_reader = open(mean_conversions_path, 'r')
        mean_conversions = mean_conversions_reader.xreadlines()
        for mc in mean_conversions:
            self.mean_conversion_list.append(mc)

    def load_sentence_cuttings(self):
        sentence_cuttings_reader = open(sentence_cutting_path, 'r')
        sentence_cuttings = sentence_cuttings_reader.xreadlines()
        for sc in sentence_cuttings:
            self.sentence_cutting_list.append(sc)

    def analyze(self):
        self.load_keywords_sheet()
        self.load_bad_words()
        self.load_mean_conversions()
        self.load_sentence_cuttings()

        print self.input_file_path
        i = 0
        with open(self.input_file_path, 'rU') as input_file:
            input_reader = csv.reader(input_file, dialect=csv.excel_tab, delimiter=',')
            input_col_len = len(input_reader.next())
            for row in input_reader:
                review_text = row[12]
                review_text = review_text.strip()
                handling_text = review_text
                for cuttings in self.sentence_cutting_list:
                    if cuttings.strip() in handling_text:
                        handling_text = handling_text.replace(cuttings, '*')

                split_texts = re.split(re_pattern, handling_text)
                for sub_text in split_texts:
                    keywords_len = self.keywords_sheet.nrows - 1
                    current_row = -1
                    for kw in self.keywords_sheet.col_values(4):
                        current_row += 1
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
                            clone_row.append(self.keywords_sheet.cell_value(current_row, 2).encode('utf-8'))
                            clone_row.append(self.keywords_sheet.cell_value(current_row, 3).encode('utf-8'))
                            clone_row.append(score)
                            clone_row.append(sub_text)
                            clone_row.append(kw)
                            clone_row.append(bw)
                            self.results.append(clone_row)

    def output(self):
        self.output_file_path += self.category
        self.output_file_path += '_'
        self.output_file_path += str(time.strftime('%Y%m%d-%H%M%S'))
        self.header = ['brand', 'category', 'is_competitor', 'manufacturer', 'market', 'matched_keywords', 'online_store',
                       'product_description', 'report_date', 'retailer_product_code', 'review_date', 'review_rating',
                       'review_text', 'review_title', 'sub_category', 'time_of_publication', 'upc', 'url', 'unique_ID',
                       'review type', 'review subtype', 'review rating', 'sub_text', 'keyword', 'bad word']
        with open(self.output_file_path, 'wb') as output_file:
            output_writer = csv.writer(output_file, delimiter=",", quotechar="'", quoting=csv.QUOTE_MINIMAL)
            output_writer.writerow(self.header)
            for rs in self.results:
                output_writer.writerow(rs)
        print "Finish!!!"

    def init_personal_care(self):
        self.keyword_path = personal_care_keywords_path
        self.bad_word_path = personal_care_bad_words_path

    def init_appliances(self):
        self.keyword_path = appliances_keywords_path
        self.bad_word_path = appliances_bad_words_path

    def init_baby_care(self):
        self.keyword_path = baby_care_keywords_path
        self.bad_word_path = baby_care_bad_words_path

    def init_fabric_care(self):
        self.keyword_path = fabric_care_keywords_path
        self.bad_word_path = fabric_care_bad_words_path

    def init_feminine_care(self):
        self.keyword_path = feminine_care_keywords_path
        self.bad_word_path = feminine_care_bad_words_path

    def init_hair_care(self):
        self.keyword_path = hair_care_keywords_path
        self.bad_word_path = hair_care_bad_words_path

    def init_oral_care(self):
        self.keyword_path = oral_care_keywords_path
        self.bad_word_path = oral_care_bad_words_path

    def init_prestige(self):
        self.keyword_path = prestige_keywords_path
        self.bad_word_path = prestige_bad_words_path

    def init_shave_care(self):
        self.keyword_path = shave_care_keywords_path
        self.bad_word_path = shave_care_bad_words_path

    def init_skin_care(self):
        self.keyword_path = skin_care_keywords_path
        self.bad_word_path = skin_care_bad_words_path
