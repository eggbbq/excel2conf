from sys import flags
import os
import re
import json
import xlrd
import logging
import argparse


# excel field value type
INT = "int"
FLOAT = "float" 
DOUBLE = "double"
BOOL = "bool"
STRING = "string"

# data container
CON_LIST = 'list'
CON_DICT = 'dict'
CON_OBJECT = 'object'
CON_MATRIX = 'matrix'
CON_MATRIX_CSR = 'matrix(csr)'


class ExcelSheetInfo:
    def __init__(self):
        self.name = ''
        self.filename = ''
        self.con_type = ''
        self.data = None
        self.fields = []
        


class ExcelFieldInfo:
    def __init__(self, name, type, index, filter):
        self.name = name
        self.type = type
        self.index = index
        self.filter = filter

        self.foreign_key = None
        if self.is_foreign_key(type):
            self.foreign_key = ExcelForeignKey(type)

    
    def is_foreign_key(self, text):
        return '|' in text


class ExcelForeignKey:
    def __init__(self, expr):
        arr = expr.split('|')
        if arr[0].endswith('[]'):
            self.result_type = CON_LIST
            self.sheet_name = arr[0][:-2]
        elif arr[0].endswith('{}'):
            self.result_type = CON_DICT
            self.sheet_name = arr[0][:-2]
        else:
            self.result_type = CON_OBJECT
            self.sheet_name = arr[0]
        self.keys = arr[1].split(',')


def get_lang_type(text):
    """convert string to python type.

    Args:
        text:str

    Returns:
        int/float/bool/str
    """
    if text == INT:
        return int
    if text == FLOAT:
        return float
    if text == DOUBLE:
        return float
    if text == BOOL:
        return bool
    if text == STRING:
        return str
    return None


def get_default_value(t):
    """Get default value of the type.

    Args:
        t: int/float/bool/str

    Returns:
        The default value of the give type.
    """
    if t == int:
        return 0
    if t == float:
        return 0
    if t == bool:
        return False
    if t == str:
        return ''
    return None


def change_type(o, t):
    """Change object type.

    Args:
        o:any
        t:int/float/bool/str

    Returns:
        A value of (int/float/bool/str).
    """
    if o is None:
        return get_default_value(t)

    ot = type(o)
    if ot == t:
        return o
    if t == int:
        try:
            return int(float(o))
        except:
            return 0
    if t == float:
        try:
            f = float(o)
        except:
            f = 0
        int_f = int(f)
        return int_f if int_f == f else f
    if t == bool:
        return bool(o)
    if t == str:
        return str(o)
    msg = '{0} not implemented. Only support int/float/bool/str'.format(t)
    raise Exception(msg)


def is_excel_file(filename):
    """Determine whether the file is an excel file

    Args:
        filename:str

    Returns:
        bool
    """
    return filename.endswith('.xls') or filename.endswith('.xlsx') and not '~' in filename


def is_field_type_string(text):
    """Determine whether the text is a field type.

    Args:
        text:str

    Returns:
        bool
    """
    if text == INT or text == DOUBLE or text == FLOAT or text == BOOL or text == STRING:
        return True
    elif re.match('{0}|{1}|{2}|{3}|{4}\[\]$'.format(INT, DOUBLE, FLOAT, BOOL, STRING), text):
        return True
    elif re.match(r'[a-zA-Z0-9]+(\[\]|\{\}){0,1}\|[a-zA-Z0-9\,]+', text):
        return True
    return False


def is_basic_value_type(text):
    """Determine the text is a base type express.

    Args:
        text:str

    Returns:
        bool
    """
    return (text == INT) or (text == FLOAT) or (text == DOUBLE) or (text == BOOL) or (text == STRING)


def is_basic_value_array(text):
    """Determine the text is a list type express.

    """
    if '[]' in text:
        index = text.index('[]')
        return is_basic_value_type(text[0:index])
    return False


def parse_basic_value_array(text, element_type, split_string):
    """Parse the text to a list.
    Args:
        text:str
        element_type:int/float/bool/str
        split_string:str

    Returns:
        list
    """
    if not text:
        return []
    else:
        return [change_type(x.strip(), element_type) for x in text.split(split_string)]


def parse_excel_list(sh, info):
    """Parse excel as a list structure

    | id   | name | age   |
    | ---- | ---- | ----- |
    | int  | str  | int   |
    | c    |      |       |
    | ID   | 名称  | 年纪  |
    | 1    | Jack | 10    |
    | 2    | Lucy | 20    |
    """
    fields = []
    info.fields = fields
    start_at = 5

    for c in range(0, sh.ncols):
        fieldname = sh.cell(1, c).value
        if not fieldname:
            continue
        fieldname = str(fieldname).replace(' ', '')
        if re.match(r'[^a-zA-Z0-9_]+$', fieldname):
            continue

        field_type_string = str(sh.cell(2, c).value)
        
        if not field_type_string or not is_field_type_string(field_type_string):
            continue
        field_type_string = str(field_type_string).replace(' ', '')

        mvalue = sh.cell(3, c).value
        filter_string = filter_string = str(mvalue)
        field = ExcelFieldInfo(fieldname, field_type_string, c, filter_string)
        fields.append(field)

    info.data = []
    for r in range(start_at, sh.nrows):
        obj = {}
        info.data.append(obj)
        for field in fields:
            col = field.index
            val = sh.cell(r, col).value
            if is_basic_value_type(field.type):
                py_type = get_lang_type(field.type)
                val = change_type(val, py_type)
            elif is_basic_value_array(field.type):
                idx = field.type.index('[]')
                py_type = get_lang_type(field.type[:1-idx])
                split_string = field.type[idx+2:]
                if not split_string:
                    split_string = ','
                val = parse_basic_value_array(str(val), py_type, split_string)
            else:
                # query type
                val = str(val).split(',')

            obj[field.name] = val
    return info


def parse_excel_object(sh, info):
    """ Sheet name is 'object'.
    | id   | int  |     | ID   | 10    |
    | ---- | ---- | --- | ---- | ----- |
    | name | int  |  s  | 名称  | 20    |
    | age  | int  |     | 年纪  | 30    |
    """
    fields = []
    obj = {}
    info.fields = fields
    info.data = obj

    for r in range(1, sh.nrows):
        field_name = str(sh.cell(r, 0).value)
        field_type = str(sh.cell(r, 1).value)
        field_filter = str(sh.cell(r, 2).value)
        field_value = sh.cell(r, 4).value
        
        py_type = get_lang_type(field_type)

        if field_name and py_type:
            fields.append(ExcelFieldInfo(field_name, field_type, r, field_filter))
            val = change_type(field_value, py_type)
            obj[field_name] = val
    
    return info


def parse_excel_mat(sh, info):
    """Sheet name is matrix. CSR

    | matrix |     |     |
    | ------ | --- | --- |
    |        | 11  | 12  |    
    |        | 21  | 22  |
    """

    pyt = int
    info.fields = [
        ExcelFieldInfo('type', sh.name, 0, ''),
        ExcelFieldInfo('row_head', 'int[]', 0, ''),
        ExcelFieldInfo('col_head', 'int[]', 1, ''),
        ExcelFieldInfo('matrix', 'int[]', 2, '')
    ]
    mat = []

    if info.con_type == CON_MATRIX_CSR:
        row_count = 0
        mat.append(0)

        # 读取非0行
        col_items = []
        p = 1
        for r in range(1, sh.nrows):
            col_count = 0
            for c in range(1, sh.ncols):
                val = sh.cell(r, c).value
                val = change_type(val, pyt)
                if not val == 0:
                    col_count = col_count + 1
                    col_items.append(c - 1)
                    col_items.append(val)
            if col_count > 0:
                row_count = row_count + 1
                mat.append(r - 1)
                mat.append(p)
                p = p + col_count*2

        for i in range(1, row_count):
            index = i*2+1+1
            mat[index] = mat[index] + row_count*2

        mat[0] = row_count
        for x in col_items:
            mat.append(x)

        mat[0] = row_count
    else:
        for r in range(1, sh.nrows):
            col_count = 0
            for c in range(1, sh.ncols):
                val = sh.cell(r, c).value
                val = change_type(val, pyt)
                mat.append(val)

    col_head = [change_type(sh.cell(1, i).value, pyt) for i in range(1, sh.ncols)]
    row_head = [change_type(sh.cell(i, 0).value, pyt) for i in range(2, sh.nrows)]

    info.data = {"matrix": mat, "col_head": col_head, "row_head": row_head}

    return info


def get_container_type(book, sh):
    return str(sh.cell(0,0).value)


def filter_fields(info_dict, filter_string):
    """Remove fields which is matched in the filter string

    Args:
        info_dict:dict
        match:string
    """
    for sheet_info in info_dict.values():
        delete_keys = []
        for i in range(len(sheet_info.fields), 0, -1):
            field = sheet_info.fields[i]
            if field.filter and field.filter in filter_string:
                delete_keys.append(f.name)
                del sheet_info.fields[i]
        
        if sheet_info.con_type == CON_OBJECT:
            for k in delete_keys:
                del sheet_info.data[k]
        elif sheet_info.con_type == CON_LIST:
            for obj in sheet_info.data:
                for k in delete_keys:
                    del obj[k]
        elif sheet_info.con_type == CON_DICT:
            if type(sheet_info.data) == list:
                for obj in sheet_info.data:
                    if type(obj) == dict:
                        for k in delete_keys:
                            del obj[k]
            elif type(sheet_info.data) == dict:
                for obj in sheet_info.data.values():
                    if type(obj) == dict:
                        for k in delete_keys:
                            del obj[k]
        else:
            print('Error:matrix does not support filter')


def assemble_foreign_item(info_dict):
    """组装外链对象

    Args:
        info_dict
    """
    for sheet_info in info_dict.values():
        if sheet_info.con_type == CON_DICT or sheet_info.con_type == CON_LIST:
            foreign_key_fields = [f for f in sheet_info.fields if f.foreign_key]
            for f in foreign_key_fields:
                attrs = f.foreign_key.keys
                f_sheet_name = f.foreign_key.sheet_name
                field_name = f.name
                con_result_type = f.foreign_key.result_type

                if not f_sheet_name in info_dict:
                    print('Error:Foreign sheet [{0}] is not found'.format(f_sheet_name))
                    continue

                f_sheet_info = info_dict[f_sheet_name]
                if not(f_sheet_info.con_type == CON_LIST or f_sheet_info.con_type == CON_DICT):
                    print('Error:The con_type of the foreign sheet [{0}] must be \'list\' or \'dict\''.format(f_sheet_name))
                    continue

                for item in sheet_info.data:
                    conds = item[field_name]
                    if not conds:
                        continue

                    foreign_result = None
                    for fobj in f_sheet_info.data:
                        found = True
                        for (cond, attr) in zip(conds, attrs):
                            val = fobj[attr]
                            if not (change_type(cond, type(val)) == val):
                                found = False
                                break
                        
                        if found:
                            if con_result_type == CON_LIST:
                                if foreign_result is None:
                                    foreign_result = []
                                foreign_result.append(fobj)
                            elif con_result_type == CON_DICT:
                                if foreign_result is None:
                                    foreign_result = {}
                                pk = f_sheet_info.fields[0].name
                                k = fobj[pk]
                                foreign_result[k] = fobj
                            elif con_result_type == CON_OBJECT:
                                foreign_result = fobj
                                break
                            
                    if foreign_result:
                        item[field_name] = foreign_result
                    else:
                        print("Error:forign not found=>{0}.{1} {2}:{3} {4}".format(sheet_info.filename,sheet_info.name, field_name, attrs, conds))


                



def assemble_data_dict(info_dict):
    data_dict = {}
    for sheet_info in info_dict.values():
        if sheet_info.con_type == CON_DICT:
            key_field_name = sheet_info.fields[0].name
            data = {}
            for item in sheet_info.data:
                key = item[key_field_name]
                data[key] = item
        else:
            data = sheet_info.data
        data_dict[sheet_info.name] = data
    return data_dict


def merge_array_item_fields(info_dict):
    """如果字段中出现下划线加数字结尾，则视为数组结构，进行合并操作
    ids_0:int,ids_1:int... => ids:int[]

    Args:
        info_dict
    """
    for sheet_info in info_dict.values():
        # find array item mark
        merge_fields_dict = {}
        for f in sheet_info.fields:
            split_pos = f.name.rfind('_')
            field_name = f.name[:split_pos]
            if split_pos > 0 and f.name[split_pos+1:].isnumeric():
                is_exists = False
                for f2 in sheet_info.fields:
                    if f2.name == field_name:
                        is_exists
                        break
                
                if not is_exists:
                    if not field_name in merge_fields_dict:
                        merge_fields_dict[field_name] = []
                    merge_fields_dict[field_name].append(f)
            else:
                # not array item
                pass
        

        for new_field_name in merge_fields_dict:
            merge_fields = merge_fields_dict[new_field_name]
            sheet_info.fields.append(ExcelFieldInfo(new_field_name, merge_fields[0].type+"[]", -1, merge_fields[0].filter))
            
            # delete fields
            for f in merge_fields:
                sheet_info.fields.remove(f)
            
            # merge data
            if sheet_info.con_type == CON_LIST or sheet_info.con_type == CON_DICT:
                for obj in sheet_info.data:
                    lst = []
                    for f in merge_fields:
                        lst.append(obj[f.name])
                        del obj[f.name]
                    obj[new_field_name] = lst
            elif sheet_info.con_type == CON_OBJECT:
                lst = []
                obj = sheet_info.data
                for f in merge_fields:
                    lst.append(obj[f.name])
                    del obj[f.name]
                obj[new_field_name] = lst


def assemble_simple_array_sheet(info_dict):
    """构造简单的数组

    Args:
        info_dict
    """
    for sheet_info in info_dict.values():
        key = sheet_info.fields[0].name
        if sheet_info.con_type == CON_LIST and 1 == len(sheet_info.fields) and key == '_':
            lst = [x[key] for x in sheet_info.data]
            sheet_info.data = lst


def assemble_meta_dict(info_dict):
    """生成meta表

    Args:
        info_dict

    Returns:
        dict
    """
    meta_dict = {}
    for sheet_info in info_dict.values():
        fields = [{"name":x.name, "type":x.type} for x in sheet_info.fields]        
        for f in fields:
            if '|' in f["type"]:
                f["type"] = f["type"].split('|')[0]
        pk = sheet_info.fields[0].name if sheet_info.con_type == CON_DICT else ''
        type = sheet_info.con_type
        filename = sheet_info.filename

        meta_dict[sheet_info.name] = {"type":type, 'filename':filename, 'name':sheet_info.name, 'fields':fields, 'primary_key':pk}
    return meta_dict


def diff_meta(meta_filepath, meta):
    """比较当前meta信息和历史记录的meta信息是否不同

    Args:
        last_meta_file_path
        meta
    
    Returns:
        Ture meta信息变更
        False meta信息没有变更
    """
    if os.path.isfile(meta_filepath):
        with open(meta_filepath, encoding='utf-8') as f:
            last_def = json.load(f)
            lst  = []
            for k in meta:
                if k in last_def and not meta[k] == last_def[k]:
                    lst.append('{0}.{1}'.format(meta[k]['filename'], k))
            return lst
    return False


def get_excels_info_dict(excel_dir, ignore_filenames):
    """读取目录下的Excel文件转换成预处理的数据结构

    Args:
        excel_dir
        ignore_filenames
    
    Returns:
        dict
    """
    ignore_filenames = [x.lower() for x in ignore_filenames]
    info_dict = {}
    for filename in os.listdir(excel_dir):
        if not is_excel_file(filename) or re.search(r'[^a-zA-Z0-9_+\-.]', filename) or (filename.lower() in ignore_filenames):
            continue
        filepath = os.path.join(excel_dir, filename)
        book = xlrd.open_workbook(filepath, encoding_override='utf-8')
        
        filename_no_ext = os.path.splitext(filename)[0]
        filename_no_ext = filename_no_ext.replace('+', '').replace('-', '')

        for sh in book.sheets():
            if sh.nrows == 0:
                continue
            con_type  = get_container_type(book, sh)
            sheet_info = ExcelSheetInfo()
            sheet_info.filename = filename_no_ext
            sheet_info.con_type = con_type
            sheet_info.name = sh.name

            if con_type == CON_LIST:
                if sh.nrows >= 5:
                    parse_excel_list(sh, sheet_info)
            elif con_type == CON_DICT:
                if sh.nrows >= 5:
                    parse_excel_list(sh, sheet_info)
            elif con_type == CON_OBJECT:
                if sh.ncols >= 5:
                    parse_excel_object(sh, sheet_info)
            elif con_type == CON_MATRIX or con_type == CON_MATRIX_CSR:
                if sh.nrows >= 2:
                    parse_excel_mat(sh, sheet_info)

            if sheet_info.fields:
                if sh.name in info_dict:
                    info = info_dict[sh.name]
                    print("Error: {0}.{1} = {2}.{3}".format(info.filename, info.name, sheet_info.filename, sheet_info.name))
                else:
                    info_dict[filename_no_ext] = sheet_info
    
    return info_dict

def parse(excel_dir, filter_string, ignore_filenames):
    info_dict = get_excels_info_dict(excel_dir, ignore_filenames)

    assemble_foreign_item(info_dict)

    if filter_string:
        filter_fields(info_dict, filter_string)

    merge_array_item_fields(info_dict)
    assemble_simple_array_sheet(info_dict)

    ret_data_dict = assemble_data_dict(info_dict)
    ret_meta_dict = assemble_meta_dict(info_dict)

    # del foreign sheet
    for sheet_info in info_dict.values():
        for f in sheet_info.fields:
            if f.foreign_key:
                del ret_data_dict[f.foreign_key.sheet_name]

    return (ret_data_dict, ret_meta_dict)


def main():
    args = argparse.ArgumentParser()
    args.add_argument('--excel_dir', default='./')
    args.add_argument('--export_dir', default='./')
    args.add_argument('--filter', default='', help='Filter string. "c,s"')
    args.add_argument('--ignore', default='', help='Ignore list. "Item,Goods"')
    args.add_argument('--merge_to_file', default='config.json', help='It is used when iseparate_type=3 ')
    args.add_argument('--separate_type', default= 3, type=int, help="1 separate with sheet, 2 separate with file  3 all in one")
    args.add_argument('--chdir', default=None)
    args.add_argument('--param', default=None, help='init argument file')
    arg = args.parse_args()

    if arg.param:
        if os.path.exists(arg.param):
            with open(arg.param, mode='r', encoding='utf-8') as f:
                param = json.load(f)
                chdir = param.get('chdir', None)
                ignore = param.get('ignore', '')
                filter = param.get('filter', '')
                excel_dir = param.get('excel_dir', './')            
                export_dir = param.get('export_dir', './')
                merge_to_file = param.get('merge_to_file', 'config.json')
                separate_type = param.get('separate_type', 3)
        else:
            param = {                
                'excel_dir':'./',
                'export_dir':'./',
                'merge_to_file':'config.json',
                'separate_type':3,
                'ignore':'',
                'filter':'',
            }
            with open(arg.param, mode='w') as f:
                json.dump(param, f, ensure_ascii=False, indent=True)
            print('{0} is created.... first time. This time i did not read excel files'.format(arg.param))
            return
            
    else:
        chdir = arg.chdir
        ignore = arg.ignore
        filter = arg.filter
        excel_dir = arg.excel_dir        
        export_dir = arg.export_dir
        merge_to_file = arg.merge_to_file
        separate_type = arg.separate_type

    if chdir:
        os.chdir(chdir)
    
    ignore_filenames = ignore.split(',')
    data, meta = parse(excel_dir, filter, ignore_filenames)

    meta_filepath = os.path.join(excel_dir,'.meta.txt')
    changed_items = diff_meta(meta_filepath, meta)

    if diff_meta(meta_filepath, meta):
        print("Error:")
        print('Those are excel files which field defines are changed.')
        print(changed_items)
        print('Please check out carefully. \nIf you make sure to create new meta file, please delete \".meta.txt\" at first.\n')
    else:
        with open(meta_filepath, mode='w') as f:
            json.dump(meta, f, ensure_ascii=False, indent=True)

        if separate_type == 3:
            # all in one
            json_filepath = os.path.join(export_dir, merge_to_file)
            with open(json_filepath, mode='w') as f:
                json.dump(data, f, ensure_ascii=False)
        elif separate_type == 2:
            # Separate with excel file
            group = {}
            for k in meta:
                m = meta[k]
                filename = m['filename']
                if not filename in group:
                    group[filename] = []
                group[filename].append(m['name'])
            
            for k in group:
                file_pack = {}
                for sheet_name in group[k]:
                    if sheet_name in data:
                        file_pack[sheet_name] = data[sheet_name]
                
                if file_pack:
                    json_filepath = os.path.join(export_dir, k+'.json')
                    with open(json_filepath, mode='w') as f:
                        json.dump(file_pack, f, ensure_ascii=False)


        elif separate_type == 1:
            # Separate with sheet
            for k in data:
                sheet_pack = data[k]
                json_filepath = os.path.join(export_dir, k+'.json')
                with open(json_filepath, mode='w') as f:
                    json.dump(sheet_pack, f, ensure_ascii=False)
        else:
            print("Error:separate_type value error.")




if __name__ == '__main__':
    main()



        



        

            

    
    