import io
import os
import sys
from textwrap import dedent, indent

import pandas as pd
import numpy as np
from slugify import slugify #python-slugify

from rdmo_client import Client

# HiddenPrints-class taken from https://stackoverflow.com/a/45669280/4649719 
# (licensed under CC-BY-SA 4.0; (c) Alexander C [stackoverflow-username])
class HiddenPrints:
    def __enter__(self):
        self._original_stdout = sys.stdout
        sys.stdout = open(os.devnull, 'w')

    def __exit__(self, exc_type, exc_val, exc_tb):
        sys.stdout.close()
        sys.stdout = self._original_stdout

class xlsx2rdmo_lite:

    def __init__(self):
        pass

    def import_to_rdmo(self, xlsx_path):
        self._read_xlsx(xlsx_path)
        self._create_catalog()
        self._create_sections_and_pages()
        self._create_questionsets()
        self._create_questions()

    def init_rdmo_access(self, base_url, auth=('admin','admin'), token=None, uri_prefix=None):
        self.base_url = base_url
        if uri_prefix is None:
            self.uri_prefix = base_url + '/instance'
        if not token is None: #preferring token over basic auth
            self.token = token #admintoken!
            self.client = Client(base_url, token=self.token)
        elif auth:
            self.auth = auth
            self.client = Client(base_url, auth=self.auth)

    def _read_xlsx(self, xlsx_path):
        df_from_excel = pd.read_excel(xlsx_path).replace(np.NaN, '')
        self.df_from_excel = df_from_excel.set_index([0,1,2,3])
        return self.df_from_excel
    
    def _create_catalog(self):
        title_catalog = self.df_from_excel.index.get_level_values(0).unique().item()
        catalog_key = 'catalog-'+slugify(title_catalog)
        catalog_obj = {
            "uri_prefix": self.uri_prefix,
            "uri_path": catalog_key,
            'title_de': title_catalog,
            'title_en': title_catalog
        }
        try:
            with HiddenPrints():
                self.catalog = self.client.create_catalog(
                    catalog_obj
                )
        except:
            with HiddenPrints():
                self.catalog = self.client.update_catalog(
                    [x for x in self.client.list_catalogs() if x['uri_path']==catalog_key][0]['id'],
                    catalog_obj
                )
        return self.catalog
        
    def display(self, obj):
        try:
            display(obj)
        except:
            print(obj)
            
    def _create_sections_and_pages(self):
        #also includes creation of corresponding attributes
        sections = list(self.df_from_excel.index.get_level_values(1).unique())
        result = None
        for section in sections:
            attrib_obj = {
                "uri_prefix": self.uri_prefix,
                "key": slugify(section),
            }
            try:
                with HiddenPrints():
                    attribute = self.client.create_attribute(
                        attrib_obj
                    )
            except Exception as e:
                with HiddenPrints():
                    attribute = self.client.update_attribute(
                        [x for x in self.client.list_attributes() if x['key'] == slugify(section)][0]['id'],
                        attrib_obj
                    )
            # self.display(attribute)
            section_obj = {
                "uri_prefix": self.uri_prefix,
                "uri_path": slugify(section),
                "catalogs": [
                    self.catalog["id"]
                ],
                "title_en": section,
                "title_de": section
            }
            try:
                with HiddenPrints():
                    result = self.client.create_section(
                        section_obj
                    )
            except:
                with HiddenPrints():
                    result = self.client.update_section(
                        [x for x in self.client.list_sections() if x["uri_path"]==slugify(section)][0]['id'],
                        section_obj
                    )
            # self.display(result)

            page_obj = {
                "uri_prefix": self.uri_prefix,
                "uri_path": 'page-' + slugify(section),
                "sections": [
                    [x for x in self.client.list_sections() if x["uri_path"]==slugify(section)][0]['id']
                ],
                "title_en": section,
                "title_de": section
            }
            try:
                with HiddenPrints():
                    result = self.client.create_page(
                        page_obj
                    )
            except:
                with HiddenPrints():
                    result = self.client.update_page(
                        [x for x in self.client.list_pages() if x["uri_path"]=="page-"+slugify(section)][0]['id'],
                        page_obj
                    )
            # self.display(result)

    def _create_questionsets(self):
        
        results = {}
        questionsets = {
            x: None 
            for x in zip(
                self.df_from_excel.index.get_level_values(1),
                self.df_from_excel.index.get_level_values(2)
            )
        } #hack for uniquing two cols.
        
        for i, ((section, questionset), _none) in enumerate(questionsets.items()):
            print('Questionset ' + str(i+1) + ' of ' + str(len(questionsets.keys())))
            attrib_obj = {
                "uri_prefix": self.uri_prefix,
                "key": slugify(section)+'_'+slugify(questionset),
                "parent": [x for x in self.client.list_attributes() if x['key'] == slugify(section)][0]['id']
            }
            try:
                with HiddenPrints():
                    attribute = self.client.create_attribute(
                        attrib_obj
                    )
                print('  created attribute: ' + str(slugify(section)+'_'+slugify(questionset))+ ' (ID:'+str(result['id'])+')')
            except Exception as e:
                with HiddenPrints():
                    attribute = self.client.update_attribute(
                        [x for x in self.client.list_attributes() if x['key'] ==  slugify(section)+'_'+slugify(questionset)][0]['id'],
                        attrib_obj
                    )
                print('  updated attribute: ' + str(slugify(section)+'_'+slugify(questionset))+ ' (ID:'+str(attribute['id'])+')')
            # display(attribute)
        
            relevant_pages = [x for x in self.client.list_pages() if x['uri_path'] == "page-"+slugify(section)]
            #display(relevant_page)
            questionset_name = str(slugify(section)+'_'+slugify(questionset, max_length=100))
            questionset_obj={
                "uri_prefix": self.uri_prefix,
                "uri_path": questionset_name,
                "pages": [
                  relevant_pages[0]['id'] # sadly this doesn't work!
                ],
                "title_en": questionset,
                "title_de": questionset
            }
            try:
                with HiddenPrints():
                    result = self.client.create_questionset(questionset_obj)
                print('  updated questionset: ' + questionset_name + ' (ID:'+str(result['id'])+')')
            except KeyboardInterrupt:
                raise
            except Exception as e:
                with HiddenPrints():
                    result = self.client.update_questionset(
                        [x for x in self.client.list_questionsets() if x['uri_path']==questionset_name][0]['id'],
                        questionset_obj
                    )
                print('  updated questionset: ' + questionset_name + ' (ID:'+str(result['id'])+')')
            #         display(questionset_obj)
            # display(result)
            try:
                results[str(relevant_pages[0]['id'])][0]
            except:
                results[str(relevant_pages[0]['id'])] = []
                
            result['custom_section']=section
            results[str(relevant_pages[0]['id'])].append(result)
        return results

    def _create_questions(self):
        
        for i, (row, question_series) in enumerate(
            self.df_from_excel[self.df_from_excel['widgettype']=="text"].reset_index().iterrows()
        ):
            print('Question ' + str(i+1) + ' of ' + str(self.df_from_excel.index.size))
        
            attrib_name = (slugify(question_series[1])+'_'+slugify(question_series[2], max_length=100)+'_'+slugify(question_series['frage_de'], max_length=100))[:110]
            attrib_obj = {
                "uri_prefix": self.uri_prefix,
                "key": attrib_name,
                "parent": [
                    x for x 
                    in self.client.list_attributes() #
                    if x['key'] == slugify(question_series[1])+'_'+slugify(question_series[2], max_length=100)
                ][0]['id']
            }
            try:
                with HiddenPrints():
                    attribute = self.client.create_attribute(
                        attrib_obj
                    )
                print('  created attribute (ID: ' + str(attribute['id']) + ')')
            except Exception as e:
                with HiddenPrints():
                    attribute = self.client.update_attribute(
                        [x for x in self.client.list_attributes() if x['key'] == attrib_name][0]['id'],
                        attrib_obj
                    )
                print('  updated attribute (ID: ' + str(attribute['id']) + ')')
            # display(attribute)
            
            # display(question_series.to_frame())

            question_name = 'question-'+attrib_name
            question_obj = {
                "uri_prefix": self.uri_prefix,
                "uri_path": question_name,
                "comment": question_series['comment'],
                "questionsets": [
                    [x for x in self.client.list_questionsets() if x["uri_path"] == slugify(question_series[1])+'_'+slugify(question_series[2], max_length=100)][0]['id']
                ],
                "attribute": [x for x in self.client.list_attributes() if x['key'] == attrib_name][0]['id'],
                "text_en": question_series['frage_en'],
                "default_text_en": question_series['defaultanswer_en'],
                "text_de": question_series['frage_de'],
                "default_text_de": question_series['defaultanswer_de'],
                "value_type": 'text',
                "widget_type": question_series['widgettype']
            }
            try:
                with HiddenPrints():
                    result = self.client.create_question(
                        question_obj
                    )
                print('  created question (ID: ' + str(result['id']) + ')')
            except Exception as e:
                with HiddenPrints():
                    result = self.client.update_question(
                        [x for x in self.client.list_questions() if x["uri_path"] == question_name][0]['id'],
                        question_obj
                    )
                print('  updated question (ID: ' + str(result['id']) + ')')
            # display(result)
        