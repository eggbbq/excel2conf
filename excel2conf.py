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

TYPE_TABLE = 'table'
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
        return int(float(o))
    if t == float:
        f = float(o)
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

        self.is_query_expr = '|' in filed_type
        self.query_fields = None
        self.query_table = None
        self.query_result_type = None

        if self.is_query_expr:
            arr = filed_type.split('|')
            if arr[0].endswith('[]'):
                self.query_table = arr[0][:-2]
                self.query_result_type = TYPE_LIST
                self.query_fields = arr[1].split(',')
            elif arr[0].endswith(r'{}'):
                self.query_table = arr[0][:-2]
                self.query_result_type = TYPE_DICT
                self.query_fields = arr[1].split(',')
            else:
                self.query_table = arr[0]
                self.query_result_type = TYPE_OBJECT
                self.query_fields = arr[1].split(',')


class ExcelInfo(object):
    def __init__(self):
        self.name = ''
        self.filename = ''
        self.type = None
        self.data = None
        self.fields = None
        self.start = 2


def parse_excel_list(sh, info):
    """Sheet name is list.

    | id   | name | age   |
    | ---- | ---- | ----- |
    | int  | str  | int   |
    | 1    | abc  | 20    |
    """
    fields = []
    info.fields = fields
    startrow = 2
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
                startrow = 3
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


def parse_excel_dict(sh, info):
    """Sheet name is dict.

    | id   | name | age   |
    | ---- | ---- | ----- |
    | int  | str  | int   |
    | 1    | abc  | 20    |
    """
    parse_excel_list(sh, info)
    if info.data:
        dic = {}
        key_name = info.fields[0].name
        for x in info.data:
            key = x[key_name]
            dic[key] = x
        info.data = dic
    return info


def parse_excel_table(sh, info):
    """Sheet name is table.

    | int | v1  | v2  |
    | --- | --- | --- |
    | id  | 11  | 12  |    
    | nam | 21  | 22  |
    """
    cell0 = str(sh.cell(0, 0).value)
    pyt = pytype(cell0)
    head = []
    array = []
    info.fields = [FieldInfo('', cell0, -1, '')]
    info.data = {
        'head': head,
        'array': array
    }
    for c in range(1, sh.ncols):
        val = sh.cell(0, c).value
        head.append(val)
    for r in range(1, sh.nrows):
        for c in range(1, sh.ncols):
            val = sh.cell(r, c).value
            val = convert(val, pyt)
            array.append(val)
    return info


def parse_excel_object(sh, info):
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

    for r in range(0, sh.nrows):
        if str(sh.cell(r, 2).value).startswith('*'):
            valcol = 3
            break

    for r in range(0, sh.nrows):
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
    """Read excel sheet name as the export type(list/dict/object/table).
    if the excel sheet name is not any of (list/dict/object/table), use list instead.

    Args:
        sh:xlrd.sheet

    Returns:
        list/dict/object/table
    """
    sheet_name = sh.name.lower()
    if sheet_name == TYPE_OBJECT:
        return TYPE_OBJECT
    elif sheet_name == TYPE_TABLE:
        return TYPE_TABLE
    return TYPE_LIST


def parse_excels(src, match, excludes):
    """Parse all excels in the src folder.

    Args:
        src:excel folder
        match:filter partten
        excludes:ignore excel files

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
            parse_excel_dict(sh, info)
        if info.type == TYPE_LIST:
            parse_excel_list(sh, info)
        if info.type == TYPE_OBJECT:
            parse_excel_object(sh, info)
        if info.type == TYPE_TABLE:
            parse_excel_table(sh, info)

        infos[name] = info

    for info in infos.values():
        if info.type == TYPE_LIST or info.type == TYPE_DICT:
            # process query
            query_fields = [t for t in info.fields if t.is_query_expr]
            for qf in query_fields:
                attrs = qf.query_fields
                table = qf.query_table
                field_name = qf.name
                query_result_type = qf.query_result_type
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

    dic = {}
    message = {}
    for info in infos.values():
        msg = {}
        for t in info.fields:
            if t.is_query_expr:
                msg[t.name] = t.type.split('|')[0]
            else:
                msg[t.name] = t.type
        message[info.name] = {"type": info.type, "message": msg}
        if info.type == TYPE_DICT:
            key_name = info.fields[0].name
            data = {}
            for t in info.data:
                key = t[key_name]
                data[key] = t
            info.data = data
        dic[info.name] = info.data

    return (dic, message)


def main():
    args = argparse.ArgumentParser()
    args.add_argument('--excel', default='./', help='excel files directory')
    args.add_argument('--match', default='', help='+? -? ? -* *')
    args.add_argument('--file',  default='configs.json')
    args.add_argument('--excludes', default='')
    arg = args.parse_args()

    excludes = arg.excludes.split(',')

    dic, msg = parse_excels(arg.excel, arg.match, excludes)
    with open(arg.file, mode='w', encoding='utf-8') as f:
        text = json.dumps(dic, ensure_ascii=False)
        f.write(text)

    with open(os.path.join(arg.excel, '__message__.json'), mode='w') as f:
        text = json.dumps(msg, ensure_ascii=False)
        f.write(text)


if __name__ == '__main__':
    main()
