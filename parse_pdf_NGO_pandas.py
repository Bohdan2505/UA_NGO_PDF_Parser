from glob import glob
import re
import os

from typing import Container
from io import BytesIO
from pdfminer.pdfinterp import PDFResourceManager, PDFPageInterpreter
from pdfminer.converter import TextConverter, XMLConverter, HTMLConverter
from pdfminer.layout import LAParams
from pdfminer.pdfpage import PDFPage

from pandas import read_excel, DataFrame, to_numeric, isnull

from qgis.PyQt.QtCore import QVariant

from qgis.PyQt.QtWidgets import QMessageBox

from qgis.core import QgsProject, QgsVectorLayer, QgsFeature, QgsFields, QgsField


df_table = DataFrame(columns=['ngo_number', 'ngo_value', 'cadnum', 'file_name', 'date'])
df_table.fillna('', inplace=True) 

def convert_pdf(
    path: str,
    file_format: str = "txt",
    write_file: bool = False,
    codec: str = "utf-8",
    password: str = "",
    maxpages: int = 0,
    caching: bool = True,
    pagenos: Container[int] = set(),
) -> str:
    try:

        rsrcmgr = PDFResourceManager()
        retstr = BytesIO()
        laparams = LAParams()
        if file_format == "txt":
            device = TextConverter(rsrcmgr, retstr, codec=codec, laparams=laparams)
        elif file_format == "html":
            device = HTMLConverter(rsrcmgr, retstr, codec=codec, laparams=laparams)
        elif file_format == "xml":
            device = XMLConverter(rsrcmgr, retstr, codec=codec, laparams=laparams)
        else:
            raise ValueError("provide file_format, either text, html or xml!")
        fp = open(path, "rb")
        interpreter = PDFPageInterpreter(rsrcmgr, device)
        for page in PDFPage.get_pages(
            fp,
            pagenos,
            maxpages=maxpages,
            password=password,
            caching=caching,
            check_extractable=True,
        ):
            interpreter.process_page(page)
        text = retstr.getvalue().decode()
        fp.close()
        device.close()
        retstr.close()
        
        name = path.split('/')[-1].split('.pdf')[0]
        #print(name)
        if write_file:
            with open(f'/{name}.{file_format}', 'w') as f:
                f.write(text)
#         print(f'ALL IS OK: {path}')
        return text
    except Exception as e:
        print(path)
        print(str(e))


def create_result_layer(ngo_data_list):
    fields = [
    QgsField('date', QVariant.String),
    QgsField('ngo_value', QVariant.Double),
    QgsField('cadnum', QVariant.String),
    QgsField('file_name', QVariant.String),
    QgsField('ngo_number', QVariant.String),
    ]
    
    for field in fields:
        if field.name() == 'date': field.setAlias("Дата витягу"),
        if field.name() == 'ngo_value': field.setAlias("НГО ділянки, грн")
        if field.name() == 'cadnum': field.setAlias("Кадастровий номер")
        if field.name() == 'file_name': field.setAlias("Шлях до витягу")
        if field.name() == 'ngo_number': field.setAlias("Номер витягу")
    
    
    # print(QVariant.Date)
    # print(QgsField('date', QVariant.Date).setAlias("Дата витягу"))
    # print(fields)
    temp_layer = QgsVectorLayer("No Geometry", 'Таблиця з відомостями НГО', 'memory')
    temp_layer_pr = temp_layer.dataProvider()
    temp_layer.startEditing()
    temp_layer_pr.addAttributes(fields)
    temp_layer.updateFields() 

    
    features_list = []
    
    
    for ngo_data_dict in ngo_data_list:
        feature = QgsFeature()
        feature.setFields(temp_layer.fields())
               
        feature['date'] = ngo_data_dict['date']
        feature['ngo_value'] = float(ngo_data_dict['ngo_value'])
        feature['cadnum'] = ngo_data_dict['cadnum']
        feature['file_name'] = ngo_data_dict['file_name']
        feature['ngo_number'] = ngo_data_dict['ngo_number']
        
        features_list.append(feature)
    

    temp_layer.addFeatures(features_list)
    # temp_layer.commitChanges()
    QgsProject.instance().addMapLayer(temp_layer)
    temp_layer.commitChanges()
    


def parse_pdf_files(path_to_folder, path_to_excel, is_add_layer):

    path_to_pdf_folder = os.path.join(path_to_folder, '*')

    directories_list = glob(rf"{path_to_pdf_folder}.pdf", recursive = True)

    general_ngo_data_list = []

    for d in directories_list:
        row_number = None
        row_number = directories_list.index(d)
        result = convert_pdf(d, 'txt', False)
        ngo_data_dict = {}
        for x in str.splitlines(result):

            cadnum = None
            date = None
            ngo_value = None
            ngo_number = None
            cadnum_check = re.match("^([0-9]{10}:[0-9]{2}:[0-9]{3}:[0-9]{4})+$", x)
            ngo_value_check = re.match("^([0-9]+.[0-9]+)+$", x)
            date_check = re.match("^([0-9]{2}.[0-9]{2}.[0-9]{4})+$", x)
            ngo_number_check = re.match("Витяг.*№.*([0-9]+)+$", x)

            if ngo_number_check:
                ngo_number = str(ngo_number_check.group()).split("№ ")
                if len(ngo_number) > 1:
                    ngo_data_dict['ngo_number'] = ngo_number[1]
                else:
                    pass
            if date_check:
                ngo_data_dict['date'] = date_check.group()
            if ngo_value_check and not date_check:
                ngo_data_dict['ngo_value'] = ngo_value_check.group()
            if cadnum_check:
                ngo_data_dict['cadnum'] = cadnum_check.group()
            ngo_data_dict['file_name'] = d
        
        if 'date' in ngo_data_dict.keys() and 'ngo_value' in ngo_data_dict.keys():
            if is_add_layer:
                general_ngo_data_list.append(ngo_data_dict)
            df_table.at[row_number, 'date'] = ngo_data_dict['date']
            df_table.at[row_number, 'ngo_value'] = float(ngo_data_dict['ngo_value'])
            df_table.at[row_number, 'cadnum'] = ngo_data_dict['cadnum']
            df_table.at[row_number, 'file_name'] = ngo_data_dict['file_name']
            df_table.at[row_number, 'ngo_number'] = ngo_data_dict['ngo_number']
    
    if is_add_layer:
         create_result_layer(general_ngo_data_list)

    df_table['file_name'] = df_table['file_name'].apply(lambda x: f'=HYPERLINK("{os.path.abspath(x)}", "{x}")')
        
    df_table.to_excel(path_to_excel, index=False) 
    QMessageBox(QMessageBox.Information, 'Процес завершено', f'Зчитування витягів НГО завершено. Результат: \n Всього файлів pdf у папці:{len(directories_list)}. \n Кількість витягів з яких успішно зчитано інформацію: {df_table.shape[0]}. \n Дякую, що використали цей плагін!', QMessageBox.Ok).exec_()
    
    
