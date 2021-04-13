# -*- coding: utf-8 -*-

import os
import re
import json
import xlrd
import argparse


TYPE_INT = 'int'
TYPE_FLOAT = 'float'
TYPE_DOUBBLE = 'doubble'
TYPE_BOOL = 'bool'
TYPE_STRING = 'string'

TYPE_MATRIX = 'matrix'
TYPE_LIST = 'list'
TYPE_DICT = 'dict'
TYPE_OBJECT = 'object'


def pytype(text):
    """convert string to python type.

    Args:
        text:str

    Returns:
        int/float/bool/str
    """
    if text == TYPE_INT:
        return int
    if text == TYPE_FLOAT:
        return float
    if text == TYPE_DOUBBLE:
        return float
    if text == TYPE_BOOL:
        return bool
    if text == TYPE_STRING:
        return str
    return None


def default(t):
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


def convert(o, t):
    """Change object type.

    Args:
        o:any
        t:int/float/bool/str

    Returns:
        A value of (int/float/bool/str).
    """
    if o is None:
        return default(t)

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


def isexcelfile(filename):
    """Determine whether the file is an excel file

    Args:
        filename:str

    Returns:
        bool
    """
    return filename.endswith('.xls') or filename.endswith('.xlsx') and not '~' in filename


def isfieldtype(text):
    """Determine whether the text is a field type.

    Args:
        text:str

    Returns:
        bool
    """
    if text == TYPE_INT or text == TYPE_DOUBBLE or text == TYPE_FLOAT or text == TYPE_BOOL or text == TYPE_STRING:
        return True
    elif re.match(r'[a-zA-Z0-9]+(\[\]|\{\}){0,1}\|[a-zA-Z0-9\,]+', text):
        return True
    return False


def is_base_type(text):
    """Determine the text is a base type express.

    Args:
        text:str

    Returns:
        bool
    """
    return (text == TYPE_INT) or (text == TYPE_FLOAT) or (text == TYPE_DOUBBLE) or (text == TYPE_BOOL) or (text == TYPE_STRING)


def is_base_list(text):
    """Determine the text is a list type express.

    """
    if '[]' in text:
        index = text.index('[]')
        return is_base_type(text[0:index])
    return False


def parse_list_text(text, element_type, split_string):
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
        return [convert(x.strip(), element_type) for x in text.split(split_string)]


class FieldInfo(object):
    def __init__(self, name, filed_type, index, match):
        assert isinstance(name, str)
        assert isinstance(index, int)
        assert isinstance(match, str)
        assert isinstance(filed_type, str)

        self.name = name
        self.type = filed_type
        self.index = index
        self.match = match

        if ForeignKey.is_forein_key(filed_type):
            self.forein_key = ForeignKey(filed_type)
        else:
            self.forein_key = None

class ForeignKey(object):
    @staticmethod
    def is_forein_key(expr):
        return '|' in expr

    def __init__(self, expr):
        arr = expr.split('|')
        if arr[0].endswith('[]'):
            self.result_type = TYPE_LIST
            self.sheet_name = arr[0][:-2]
        elif arr[0].endswith(r'{}'):
            self.result_type = TYPE_DICT
            self.sheet_name = arr[0][:-2]
        else:
            self.result_type = TYPE_OBJECT
            self.sheet_name = arr[0]
        self.keys = arr[1].split(',')


class ExcelInfo(object):
    def __init__(self):
        self.name = ''
        self.filename = ''
        self.type = None
        self.data = None
        self.fields = None
        self.start = 2


def parse_excel_list(sh, info, start):
    """Sheet name is list.

    | id   | name | age   |
    | ---- | ---- | ----- |
    | int  | str  | int   |
    | 1    | abc  | 20    |
    """
    fields = []
    info.fields = fields
    startrow = 2 + start
    for c in range(0, sh.ncols):
        fvalue = sh.cell(0, c).value
        if not fvalue:
            continue
        fieldname = str(fvalue).replace(' ', '')
        if re.match(r'[^a-zA-Z0-9_]+$', fieldname):
            continue

        tvalue = sh.cell(1, c).value
        if not tvalue or not isfieldtype(str(tvalue)):
            continue
        fieldtype = str(tvalue).replace(' ', '')

        mvalue = sh.cell(2, c).value
        mstring = ''
        if not mvalue:
            mstring = ''
        else:
            mstring = str(mvalue)
            if not mstring.startswith('*'):
                mstring = ''
            else:
                startrow = 3 + start
        field = FieldInfo(fieldname, fieldtype, c, mstring)
        fields.append(field)

    info.data = []
    for r in range(startrow, sh.nrows):
        obj = {}
        info.data.append(obj)
        for field in fields:
            col = field.index
            val = sh.cell(r, col).value
            if is_base_type(field.type):
                pyt = pytype(field.type)
                val = convert(val, pyt)
            elif is_base_list(field.type):
                idx = field.type.index('[]')
                pyt = pytype(field.type[:-idx])
                split_string = field.type[idx+2:]
                if not split_string:
                    split_string = ','
                val = parse_list_text(str(val), pyt, split_string)
            else:
                # query type
                val = str(val).split(',')

            obj[field.name] = val
    return info


# def parse_excel_dict(sh, info, start):
#     """Sheet name is dict.

#     | id   | name | age   |
#     | ---- | ---- | ----- |
#     | int  | str  | int   |
#     | 1    | abc  | 20    |
#     """
#     parse_excel_list(sh, info, start)
#     if info.data:
#         dic = {}
#         key_name = info.fields[0].name
#         for x in info.data:
#             key = x[key_name]
#             dic[key] = x
#         info.data = dic
#     return info


def parse_excel_mat(sh, info, start):
    """Sheet name is matrix. CSR

    | int |     |     |
    | --- | --- | --- |
    |     | 11  | 12  |    
    |     | 21  | 22  |
    """

    start_value_value = str(sh.cell(start, start).value)
    pyt = pytype(start_value_value)
    field_type = start_value_value+'[]'
    info.fields = [
        FieldInfo('type', sh.name, 0, ''),
        FieldInfo('row_head', field_type, 0, ''),
        FieldInfo('col_head', field_type, 1, ''),
        FieldInfo('matrix', field_type, 2, '')
    ]
    mat = []

    if sh.name.lower() == 'matrix(csr)':
        row_count = 0
        mat.append(0)

        # 读取非0行
        col_items = []
        p = 1
        for r in range(start+1, sh.nrows):
            col_count = 0        
            for c in range(start+1, sh.ncols):
                val = sh.cell(r, c).value
                val = convert(val, pyt)
                if not val == 0:
                    col_count = col_count + 1
                    col_items.append(c - start - 1)
                    col_items.append(val)
            if col_count > 0:
                row_count = row_count + 1
                mat.append(r - start - 1)
                mat.append(p)
                p = p + col_count*2
        
        for i in range(0,row_count):
            index = i*2+1+1
            mat[index] = mat[index] + row_count*2
        
        mat[0] = row_count
        for x in col_items:
            mat.append(x)
        
        mat[0] = row_count
    else:
        for r in range(start+1, sh.nrows):
            col_count = 0        
            for c in range(start+1, sh.ncols):
                val = sh.cell(r, c).value
                val = convert(val, pyt)
                mat.append(val)

    col_head = [convert(sh.cell(start, i).value, pyt) for i in range(start + 1, sh.ncols)]
    row_head = [convert(sh.cell(i, start).value, pyt) for i in range(start + 1, sh.nrows)]

    info.data = {"matrix":mat,"col_head":col_head, "row_head":row_head}

    return info


def parse_excel_object(sh, info, start):
    """ Sheet name is 'object'.

    | name | type | value |
    | ---- | ---- | ----- |
    | id   | int  | 10    |
    | name | int  | 20    |
    """
    valcol = 2
    fields = []
    obj = {}
    info.fields = fields
    info.data = obj

    for r in range(start, sh.nrows):
        if str(sh.cell(r, 2).value).startswith('*'):
            valcol = 3
            break

    for r in range(start, sh.nrows):
        fieldname = str(sh.cell(r, 0).value)
        fieldtype = str(sh.cell(r, 1).value)
        fieldvalue = sh.cell(r, valcol).value
        fieldmatch = '' if valcol == 2 else str(sh.cell(r, 2).value)
        pyt = pytype(fieldtype)

        if fieldname and pyt:
            fields.append(FieldInfo(fieldname, fieldtype, r, fieldmatch))
            val = convert(fieldvalue, pyt)
            obj[fieldname] = val

    return info


def get_export_type(sh):
    """Read excel sheet name as the export type(list/dict/object/matrix).
    if the excel sheet name is not any of (list/dict/object/matrix), use list instead.

    Args:
        sh:xlrd.sheet

    Returns:
        list/dict/object/matrix
    """
    sheet_name = sh.name.lower()
    if sheet_name == TYPE_OBJECT:
        return TYPE_OBJECT
    elif sheet_name.startswith(TYPE_MATRIX):
        return TYPE_MATRIX
    return TYPE_LIST


def parse_excels(src, match, excludes, start=0):
    """Parse all excels in the src folder.

    Args:
        src:excel folder
        match:filter partten
        excludes:ignore excel files
        start:row/col start index of data sheet

    Returns:
        (dic, msg)

    """
    excludes = [x.lower() for x in excludes]
    infos = {}
    for filename in os.listdir(src):
        if not isexcelfile(filename) or re.search(r'[^a-zA-Z0-9_+\-.]', filename) or (filename.lower() in excludes):
            continue
        filepath = os.path.join(src, filename)
        sh = xlrd.open_workbook(
            filepath, encoding_override='utf-8').sheet_by_index(0)

        name = os.path.splitext(filename)[0]
        name = name.replace('+', '').replace('-', '')

        info = ExcelInfo()
        info.name = name
        info.filename = filename
        info.type = get_export_type(sh)

        if info.type == TYPE_DICT:
            parse_excel_list(sh, info, start)
        if info.type == TYPE_LIST:
            parse_excel_list(sh, info, start)
        if info.type == TYPE_OBJECT:
            parse_excel_object(sh, info, start)
        if info.type == TYPE_MATRIX:
            parse_excel_mat(sh, info, start)

        infos[name] = info
    
    # process foreign key
    for info in infos.values():
        if info.type == TYPE_LIST or info.type == TYPE_DICT:
            query_fields = [t for t in info.fields if t.forein_key]
            for qf in query_fields:
                attrs = qf.forein_key.keys
                table = qf.forein_key.sheet_name
                field_name = qf.name
                query_result_type = qf.forein_key.result_type
                target = infos[table]

                if not (target and target.data and (target.type == TYPE_LIST or target.type == TYPE_DICT)):
                    continue

                for x in info.data:
                    conds = x[field_name]
                    if not conds:
                        continue

                    query_result = None
                    for t in target.data:
                        found = True
                        for (cond, attr) in zip(conds, attrs):
                            val = t[attr]
                            if not convert(cond, type(val)) == val:
                                found = False
                                break

                        if found:
                            if query_result_type == TYPE_LIST:
                                if query_result is None:
                                    query_result = []
                                query_result.append(t)
                            elif query_result_type == TYPE_DICT:
                                if query_result is None:
                                    query_result = {}
                                key_name = target.fields[0].name
                                key = t[key_name]
                                query_result[key] = t
                            elif query_result_type == TYPE_OBJECT:
                                query_result = t
                                break
                    x[field_name] = query_result

            # process match
            # --m +s 空和带s匹配的字段
            # --m -s 不带s匹配的字段
            # --m s  所有s匹配的字段
            # --m -* 所有无匹配的字段
            # --m *  所有有匹配的字段
            if match:
                if match == '-*':
                    allow_attrs = [t.name for t in info.fields if not t.match]
                elif match == '*':
                    allow_attrs = [t.name for t in info.fields if t.match]
                if match.startswith('-'):
                    allow_attrs = [t.name for t in info.fields if not (
                        match[1:] in t.match)]
                elif match.startswith('+'):
                    allow_attrs = [t.name for t in info.fields if not match or (
                        match[1:] in t.match)]
                else:
                    allow_attrs = [
                        t.name for t in info.fields if match in t.match]
                for x in info.data:
                    for k in list(x.keys()):
                        if not k in allow_attrs:
                            del x[k]
    
    # merge a a+b
    for info in list(infos.values()):
        if '+' in info.filename:
            del infos[info.name]
            arr = os.path.splitext(info.filename)[0].split('+')
            main_sheet = infos[arr[0]]
            if main_sheet:
                pk = main_sheet.fields[0].name
                if main_sheet.type == TYPE_LIST or main_sheet.type == TYPE_DICT and pk == info.fields[0].name:
                    for sub in info.data:
                        if not pk in sub:
                            continue
                        pk_value = sub[pk]
                        for main in main_sheet.data:
                            if pk_value == main[pk]:
                                for k in sub:
                                    main[k] = sub[k]
                                break
                # merge fields
                for i in range(1, len(info.fields)):
                    main_sheet.fields.append(info.fields[i])
    
    # merge array fields expr
    """
    {
        ids_0:1,
        ids_1:2,
        ids_3:3,
        ...
    }

    =>
    {
        ids:[1,2,3,...]
    }
    """
    for info in infos.values():
        dic = {}
        for f in info.fields:
            i = f.name.rfind('_')            
            if i > 0 and f.name[i+1:].isnumeric():
                field_name = f.name[:i]                
                if not field_name in dic:
                    dic[field_name] = f.type

                elif not dic[field_name] == f.type:
                    del dic[field_name]
                    break
        for k in dic:
            t = dic[k]
            for i in range(len(info.fields) - 1, -1, -1):
                f = info.fields[i]                
                if f.name.startswith(k):
                    del info.fields[i]

            info.fields.append(FieldInfo(k,t+"[]",-1,""))

            lst = []
            info.data            
            if info.type == TYPE_LIST or info.type == TYPE_LIST:
                for item in info.data:
                    for f in list(item.keys()):
                        if f.startswith(k):
                            lst.append(item[f])
                            del item[f]
                    item[k] = lst
            elif info.type == TYPE_OBJECT:
                obj = info.data
                for f in obj:
                    if f.startswith(k):
                        lst.append(obj[f])
                        del obj[f]
                obj[k] = lst
    # build dict
    dic = {}
    for info in infos.values():
        if info.type == TYPE_DICT:
            pk = info.fields[0].name
            data = {}
            for item in info.data:
                key = item[pk]
                data[key] = item
        else:
            data = info.data
        dic[info.name] = data
    
    # build meta
    meta = {}
    for info in infos.values():
        fields = {}
        for t in info.fields:
            if t.forein_key:
                fields[t.name] = t.type.split('|')[0]
            else:
                fields[t.name] = t.type
        if info.type == TYPE_DICT:
            pk = info.fields[0].name
        else:
            pk = ''
        meta[info.name] = {"type": info.type, "primary_key":pk, "fields": fields}

    return (dic, meta)


def main():
    args = argparse.ArgumentParser()
    args.add_argument('--excel', default='./', help='EXCEL目录')
    args.add_argument('--match', default='', help='+? -? ? -* *  字段匹配符')
    args.add_argument('--file',  default='configs.json', help='JSON导出文件')
    args.add_argument('--meta',  default='meta.txt', help='解析表格时的类型信息')
    args.add_argument('--start', default=1, type=int, help='解析表格起始行/列索引位置')
    args.add_argument('--excludes', default='')
    arg = args.parse_args()

    excludes = arg.excludes.split(',')

    dic, msg = parse_excels(arg.excel, arg.match, excludes, arg.start)
    if arg.file:
        with open(arg.file, mode='w', encoding='utf-8') as f:
            text = json.dumps(dic, ensure_ascii=False)
            f.write(text)

    if arg.meta:
        with open(os.path.join(arg.meta), mode='w') as f:
            text = json.dumps(msg, ensure_ascii=False, indent=True)
            f.write(text)


if __name__ == '__main__':
    main()
