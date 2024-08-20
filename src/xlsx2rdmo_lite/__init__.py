import io
import os
import sys
from textwrap import dedent, indent

import pandas as pd
import numpy as np
from slugify import slugify #python-slugify

from rdmo_client import Client

from requests.exceptions import HTTPError

try:
    from IPython.display import display, Markdown
except:
    display = print
    Markdown = str

# HiddenPrints-class taken from https://stackoverflow.com/a/45669280/4649719 
# (licensed under CC-BY-SA 4.0; (c) Alexander C [stackoverflow-username])
class HiddenPrints:
    def __init__(self, debug=False):
        self.debug = debug
        
    def __enter__(self):
        if self.debug==True:
            pass
        else:
            self._original_stdout = sys.stdout
            sys.stdout = open(os.devnull, 'w')

    def __exit__(self, exc_type, exc_val, exc_tb):
        if self.debug==True:
            pass
        else:
            sys.stdout.close()
            sys.stdout = self._original_stdout

class xlsx2rdmo_lite:

    def __init__(self, debug=False):
        self.debug = debug
        pass

        
    def display(self, obj):
        try:
            if (type(obj)==pd.DataFrame or type(obj)==pd.Series or type(obj)==Markdown):
                display(obj)
        except:
            print(obj)
        try:
            from pprint import pprint
            pprint(obj)
        except:
            print(obj)
            
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

    def _delete_everything_format_c(self):
        for x in self.client.list_questions():
            self.client.destroy_question(x['id'])
        for x in self.client.list_questionsets():
            self.client.destroy_questionset(x['id'])
        for x in self.client.list_sections():
            self.client.destroy_section(x['id'])
        for x in self.client.list_pages():
            self.client.destroy_page(x['id'])
        for x in self.client.list_catalogs():
            self.client.destroy_catalog(x['id'])
        for x in self.client.list_attributes():
            with HiddenPrints(self.debug):
                try: 
                    self.client.destroy_attribute(x['id']) #in try-block, because it also deletes nested attributes in a cascaded matter.
                except HTTPError: pass
            
    def _create_catalog(self):
        self.display(Markdown('### Create Catalog'))
        title_catalog = self.df_from_excel.index.get_level_values(0).unique().item()
        catalog_key = 'catalog-'+slugify(title_catalog)
        catalog_obj = {
            "uri_prefix": self.uri_prefix,
            "uri_path": catalog_key,
            'title_de': title_catalog,
            'title_en': title_catalog
        }
        try:
            with HiddenPrints(self.debug):
                self.catalog = self.client.create_catalog(
                    catalog_obj
                )
            if self.debug:
                self.display(Markdown('**catalog created** (ID: ' +str(self.catalog['id'])+ ')'))
                self.display(self.catalog)
        except:
            with HiddenPrints(self.debug):
                self.catalog = self.client.update_catalog(
                    [x for x in self.client.list_catalogs() if x['uri_path']==catalog_key][0]['id'],
                    catalog_obj
                )
            if self.debug:
                self.display(Markdown('**catalog updated** (ID: ' +str(self.catalog['id'])+ ')'))
                self.display(self.catalog)
        return self.catalog
            
    def _create_sections_and_pages(self):
        self.display(Markdown('### Create Sections and pages'))
        #also includes creation of corresponding attributes
        section_names = list(self.df_from_excel.index.get_level_values(1).unique())
        for section_name in section_names:

            ## Attribute Section Rootpoint
            section_attrib_name = slugify(
                # '{}_{}'.format(
                #     self.catalog['title'],
                #     section_name
                # )
                section_name
            )
            attrib_obj = {
                "uri_prefix": self.uri_prefix,
                "key": section_attrib_name,
            }
            with HiddenPrints(self.debug):
                try:
                    attribute = self.client.create_attribute(
                        attrib_obj
                    )
                    if self.debug:
                        self.display(Markdown('**attribute created** (ID: ' +str(attribute['id'])+ ')'))
                        self.display(attrib_obj)
                        self.display(attribute)
                except Exception as e:
                    attribute = self.client.update_attribute(
                        [x for x in self.client.list_attributes() if x['key'] == section_attrib_name][0]['id'],
                        attrib_obj
                    )
                    if self.debug:
                        self.display(Markdown('**attribute updated** (ID: ' +str(attribute['id'])+ ')'))
                        self.display(attrib_obj)
                        self.display(attribute)

            
            ## Section
            section_name_slugged = slugify(
                '{}_{}'.format(
                    self.catalog['title'],
                    section_name
                )
            )
            section_obj = {
                "uri_prefix": self.uri_prefix,
                "uri_path": section_name_slugged,
                "title_en": section_name,
                "title_de": section_name
            }
            with HiddenPrints(self.debug):
                try:
                    section = self.client.list_sections(
                        **section_obj
                    )[0]
                except IndexError:
                    try:
                        section = self.client.create_section(
                            section_obj
                        )
                    except Exception as e:
                        self.display(
                            "The requested section ({}) does not exist, and it can't be created, too. Something is wrong...".format(section_name_slugged)
                        )
                        raise e
                    else:
                        if self.debug:
                            self.display(Markdown('**section created** (ID: ' +str(section['id'])+ ')'))
                            self.display(section_obj)
                            self.display(section)
                else:
                    self.client.update_section(
                        section['id'],
                        section
                    )
                    if self.debug:
                        self.display(Markdown('**section updated** (ID: ' +str(section['id'])+ ')'))
                        self.display(section_obj)
                        self.display(section)
                if section['id'] in self.catalog['sections']:
                    self.display(Markdown(
                        '*Section has already been added to catalog (sID: {}, cID: {})'.format(section['id'], self.catalog['id'])
                    ))
                else:
                    self.catalog['sections'] = self.catalog['sections'] + [{'section': section['id'], 'order': (
                        max(
                            [f['order'] for f in self.catalog['sections']]
                        )+1 if len(
                            self.catalog['sections']
                        )>0 else 1
                    )}]
                    self.catalog = self.client.update_catalog(
                        self.catalog['id'],
                        self.catalog
                    )
                if self.debug:
                    self.display(Markdown('**catalog updated with section** (ID: ' +str(section['id'])+ ')'))
                    # self.display(section)
                    # self.display(oldcatalog)
                    # self.display(self.catalog)

            ## Page
            page_obj = {
                "uri_prefix": self.uri_prefix,
                "uri_path": 'page-' + section_name_slugged,
                "title_en": section_name,
                "title_de": section_name
            }
            with HiddenPrints(self.debug):
                try:
                    pages = self.client.list_pages(
                        **page_obj
                    )
                    if len(pages)>1:
                        raise Exception('multiple pages exist. Cant be in this limited importer...')
                    page = pages[0]
                except:
                    page = self.client.create_page(
                        page_obj
                    )
                    if self.debug:
                        self.display(Markdown('**page created** (ID: ' +str(page['id'])+ ')'))
                        self.display(page)
                if page['id'] in [f['page'] for f in section['pages']]:
                    self.display(Markdown('*Page is already added to section* (pID: {}, sID: {})'.format(page['id'], section['id'])))
                else:
                    section['pages'] = section['pages'] + [{'page': page['id'], 'order': (
                        max(
                            [f['order'] for f in section['pages']]
                        )+1 if len(
                            section['pages']
                        )>0 else 1
                    )}]
                    section = self.client.update_section(
                        section['id'],
                        section
                    )
                    if self.debug:
                        self.display(Markdown('**section updated with page (ID: {})'.format(section['id'])))
                        self.display(section)

    def _create_questionsets(self):
        self.display(Markdown('### Create Questionsets'))
        results = []
        questionsets = {
            x: None 
            for x in zip(
                self.df_from_excel.index.get_level_values(0),
                self.df_from_excel.index.get_level_values(1),
                self.df_from_excel.index.get_level_values(2)
            )
        } #hack for uniquing two cols.
        
        for i, ((catalog_name, section_name, questionset_name), _none) in enumerate(questionsets.items()):
            self.display(Markdown('*Questionset ' + str(i+1) + ' of ' + str(len(questionsets.keys()))+'*'))

            ## Attribute
            attrib_name = slugify(
                '{}_{}'.format(
                    section_name,
                    slugify(questionset_name, max_length=100)
                )
            )
            section_attrib_name = slugify(
                # '{}_{}'.format(
                #     catalog_name,
                #     section_name
                # )
                section_name
            )
            try:
                attrib_obj = {
                    "uri_prefix": self.uri_prefix,
                    "key": attrib_name,
                    "parent": [x for x in self.client.list_attributes() if x['key'] == section_attrib_name][0]['id']
                }
            except IndexError as e:
                self.display(section_attrib_name)
                raise e
            with HiddenPrints(self.debug):
                try:
                    attribute = self.client.create_attribute(
                        attrib_obj
                    )
                    self.display(Markdown('**created attribute**: ' + attrib_name + ' (ID:'+str(attribute['id'])+')'))
                except Exception as e:
                    attribute = self.client.update_attribute(
                        [x for x in self.client.list_attributes() if x['key'] == attrib_name][0]['id'],
                        attrib_obj
                    )
                self.display(Markdown('**updated attribute**: ' + attrib_name + ' (ID:'+str(attribute['id'])+')'))
            if self.debug:
                self.display(attribute)

            ## Questionset
            questionset_name_slugged = slugify(
                '{}_{}_{}'.format(
                    catalog_name,
                    section_name,
                    slugify(questionset_name, max_length=100)
                )
            )
            questionset_obj={
                "uri_prefix": self.uri_prefix,
                "uri_path": questionset_name_slugged,
                "title_en": questionset_name,
                "title_de": questionset_name
            }
            with HiddenPrints(self.debug):
                try:
                    questionset = self.client.list_questionsets(
                        **questionset_obj
                    )[0]
                except IndexError:
                    try:
                        questionset = self.client.create_questionset(questionset_obj)
                        self.display(Markdown('**created questionset**: ' + questionset_name_slugged + ' (ID:'+str(questionset['id'])+')'))
                    except Exception as e:
                        self.display(Markdown('***Something went wrong***, when creating questionset ({})'.format( questionset_name_slugged )))
                        raise e
                # except KeyboardInterrupt:
                #     raise
                
            ## Page update
            
            page_name_slugged = slugify(
                'page-{}_{}'.format(
                    self.catalog['title'],
                    section_name
                )
            )
            try:
                page = self.client.list_pages(
                    **{
                        "uri_prefix": self.uri_prefix,
                        'uri_path': page_name_slugged,
                    }
                )[0]
            except IndexError as e:
                self.display(page_name_slugged)
                self.display(
                    'Something went wrong while creating a questionset ({}), or more specific while adding the questionset to a page ({})'.format(
                        questionset_name_slugged,
                        page_name_slugged
                    )
                )
                raise e
            else:
                try:
                    if questionset['id'] in [f['questionset'] for f in page['questionsets']]:
                        self.display(Markdown('*Questionset is already added to page (qID: {}, pID: {})'.format(questionset['id'], page['id'])))
                    else:
                        page['questionsets'] = page['questionsets'] + [{'questionset': questionset['id'], 'order': (
                                    max(
                                        [f['order'] for f in page['questionsets']]
                                    )+1 if len(
                                        page['questionsets']
                                    )>0 else 1
                                )}]
                        page = self.client.update_page(
                            page['id'],
                            page
                        )
                except:
                    raise Exception('Something went wrong, when updating the page ({}) with new questionset ({})'.format(page['id'], questionset['id']))
            
                self.display(Markdown('**updated page with questionset**: ' + questionset_name_slugged + ' (ID:'+str(questionset['id'])+')'))
            if self.debug:
                self.display(questionset)
                
            questionset['custom_section']=section_attrib_name
            results.append(questionset)
        return results

    def _create_questions(self):
        self.display(Markdown('### Create Questions'))
        
        for i, (row, question_series) in enumerate(
            self.df_from_excel[self.df_from_excel['widgettype']=="text"].reset_index().iterrows()
        ):
            self.display(Markdown('*Question ' + str(i+1) + ' of ' + str(self.df_from_excel.index.size) + '*'))
        
            attrib_name = slugify(
                "{}_{}_{}".format(
                    question_series[1],
                    slugify(question_series[2], max_length=100),
                    slugify(question_series['frage_de'], max_length=100)
                )
            )[:110]
            attrib_obj = {
                "uri_prefix": self.uri_prefix,
                "key": attrib_name,
                "parent": [
                    x for x 
                    in self.client.list_attributes() #
                    if x['key'] == slugify(
                        '{}_{}'.format(
                            question_series[1],
                            slugify(question_series[2], max_length=100)
                        )
                    )
                ][0]['id']
            }
            with HiddenPrints(self.debug):
                try:
                    attribute = self.client.create_attribute(
                        attrib_obj
                    )
                    self.display(Markdown('**created attribute** (ID: ' + str(attribute['id']) + ')'))
                except Exception as e:
                    attribute = self.client.update_attribute(
                        [x for x in self.client.list_attributes() if x['key'] == attrib_name][0]['id'],
                        attrib_obj
                    )
                self.display(Markdown('**updated attribute** (ID: ' + str(attribute['id']) + ')'))
            if self.debug:
                self.display(attribute)
            
            # display(question_series.to_frame())

            question_name = 'question-'+attrib_name
            question_obj = {
                "uri_prefix": self.uri_prefix,
                "uri_path": question_name,
                "comment": question_series['comment'],
                # "questionsets": [
                #     questionset['id'] #does this work? or d I need to update the questionset?
                # ],
                "attribute": [x for x in self.client.list_attributes() if x['key'] == attrib_name][0]['id'],
                "text_en": question_series['frage_en'],
                "default_text_en": question_series['defaultanswer_en'],
                "text_de": question_series['frage_de'],
                "default_text_de": question_series['defaultanswer_de'],
                "value_type": 'text',
                "widget_type": question_series['widgettype']
            }
            with HiddenPrints(self.debug):
                questions = self.client.list_questions(
                    **{key: value for key, value in question_obj.items() if key in ["uri_prefix", "uri_path", "attribute"]} 
                )
                if len(questions)>1:
                    raise Exception('More questions found then expected (ID: {})'.format(str([f['if'] for f in questions])))
                elif len(questions)==0:
                    question = self.client.create_question(
                        question_obj
                    )
                    self.display(Markdown('**created question** (ID: {}'.format(question['id'])))
                else:
                    question = questions[0]
                    question.update(question_obj)
                    question = self.client.update_question(
                        question['id'],
                        question
                    )
                    self.display(Markdown('**updated question** (ID: ' + str(question['id']) + ')'))

            with HiddenPrints(self.debug):
                questionset_name = slugify(
                    '{}_{}_{}'.format(
                        question_series[0],
                        question_series[1],
                        slugify(question_series[2], max_length=100)
                    )
                )
                questionset = self.client.list_questionsets(
                    uri_path=questionset_name
                )[0]
                if question['id'] in [f['question'] for f in questionset['questions']]:
                    self.display('*Question has already been added to questionset* (ID: {})'.format(questionset['id']))
                else:
                    self.display(Markdown('*old questionset*'))
                    self.display(questionset)
                    questionset['questions'] = questionset['questions'] + [{
                        'question': question['id'],
                        'order': (
                            max(
                                [f['order'] for f in questionset['questions']]
                            )+1 if len(
                                questionset['questions']
                            )>0 else 1
                        )
                    }]
                    questionset = self.client.update_questionset(
                        questionset['id'],
                        questionset
                    )
                    self.display(Markdown('*modified questionset*'))
                    self.display(questionset)
                    if self.debug:
                        self.display(Markdown('**Updated Questionset with questions** (ID: {})'.format(questionset['id'])))






        