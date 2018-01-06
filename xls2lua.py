#!/usr/bin/python
# -*- coding: utf-8 -*-
#repository: https://github.com/trumanzhao/xls2lua
#trumanzhao, 2017/03/24, trumanzhao@foxmail.com

import os, sys, argparse, time, datetime, hashlib, codecs, xlrd

'''
填表时一般无需刻意对字符串加引号,除非是raw模式.
在映射名前面加'*'表示主键,主键可以有多个;如果没有主键,则会按行处理为数组,那就不能指定映射了.
映射名结尾,'?'表示bool,'#'表示数字,'@'表示字符串,如果前面三者都不是,则是raw模式.
bool类型会自动处理0,1到true,false的转换,也支持直接填true/false,是/否,有/无; EMPTY处理为false;
数字类型,如果没有填,则是0;如果填的不是数字,则抛出异常.
字符串模式,会自动加引号,EMPTY处理为"";
raw模式,转换时照搬,即可能是字符串,也可能是代码,也可能是数字或布尔,转换时字符串不会额外加引号,EMPTY处理为0;
可以考虑加个参数,使得所有EMPTY都抛异常,或者指定EMPTY咋处理.
'''

# 数字转字符串的格式,如果有不同的精度要求,可以调整这里
number2string = "%.6f";

class _ColumnDesc(object):
    """列描述"""
    def __init__(self, column_name, field_name, column_idx):
        first_char = field_name[0];
        last_char = field_name[-1];
        map_table = {u"?":"bool", u"#":"number", u"$":"string"};
        field_name = field_name if first_char != u"*" else field_name[1:];
        field_name = field_name if last_char not in map_table else field_name[:-1];
        self.column_name = column_name;
        self.column_idx = column_idx;
        self.is_key = (first_char == u"*");
        self.field_name = field_name;
        self.map_type = map_table[last_char] if last_char in map_table else "raw";

class _SheetDesc(object):
    """sheet描述"""
    def __init__(self, sheet_name, table_name):
        self.sheet_name = sheet_name;
        self.table_name = table_name;
        self.columns = list();
        self.maps = dict();
        self.keys = list();
        self.has_key = False;

    def map(self, column_name, field_name, column_idx):
        desc = _ColumnDesc(column_name, field_name, column_idx);
        self.columns.append(desc);
        self.maps[column_name] = desc;
        if desc.is_key:
            self.keys.append(desc);
            self.has_key = True;

def _unicode_anyway(text):
    try:
        some_type = unicode;
        return text.decode("utf-8") if isinstance(text, str) else text;
    except NameError:
        return text.decode("utf-8") if isinstance(text, bytes) else text;

class Converter(object):
    _scope = None;
    _indent = u"\t";
    _meta = None;
    _lines = None;
    _tables = None;

    def __init__(self, scope, indent, meta):
        self._scope = scope;
        self._indent = indent == 0 and u"\t" or u" " * indent;
        self._meta = meta;
        self.reset();

    def _get_signature(self):
        url = "https://github.com/trumanzhao/xls2lua";
        now = datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S');
        return u"--%s, %s\n" % (now, url);

    def convert(self, xls_filename):
        xls_filename = _unicode_anyway(xls_filename);
        try:
            self._workbook = xlrd.open_workbook(xls_filename);
            self._xls_filetime = os.path.getmtime(xls_filename);
            self._xls_filename = xls_filename;
        except:
            raise Exception("Failed to load workbook: %s" % xls_filename);

        self._sheet_names = self._workbook.sheet_names();
        self._meta_tables = list();

        if self._meta in self._sheet_names:
            self._load_meta_sheet();
        else:
            self._load_meta_header();

        for sheet_desc in self._meta_tables:
            self._convert_sheet(sheet_desc);
            self._tables.append(sheet_desc.table_name);

    def save(self, filename):
        lua_dir = os.path.split(filename)[0];
        if lua_dir != "" and not os.path.exists(lua_dir):
            os.makedirs(lua_dir);
        line = u"";
        if self._scope == u"local":
            for table_name in self._tables:
                if line == u"":
                    line += u"return %s" % table_name;
                else:
                    line += u", %s" % table_name;
            line += u";\n";
        code = u"".join(self._lines) + line;
        open(filename, "wb").write(code.encode("utf-8"));

    def reset(self):
        self._lines = list();
        self._lines.append(self._get_signature());
        self._lines.append(u"\n");
        self._tables = list();

    #比较文件时间戳,如果input比较新或者output不存在,则返回True,否则False
    def compare_time(self, input_file, output_file):
        if not os.path.isfile(output_file):
            return True;
        input_time = os.path.getmtime(input_file);
        output_time = os.path.getmtime(output_file);
        return input_time >= output_time;

    #meta_tables: list of _SheetDesc
    #meta_tables之所以是一个list而不是dict,是因为允许对同一个sheet做多个映射转换
    def _load_meta_sheet(self):
        meta_sheet = self._workbook.sheet_by_name(self._meta);
        for column_idx in range(0, meta_sheet.ncols):
            self._load_meta_column(meta_sheet, column_idx);

    #meta_sheet中,每列定义了一个sheet的映射
    #本函数将每列数据load为一个meta_table:
    def _load_meta_column(self, meta_sheet, column_idx):
        text = meta_sheet.cell(0, column_idx).value;
        text_split = text.split("=");
        sheet_name = text_split[0];
        table_name = text_split[1];
        if sheet_name not in self._sheet_names:
            raise Exception("Meta error, sheet not exist: %s" % sheet_name);

        data_sheet = self._workbook.sheet_by_name(sheet_name);
        column_headers = dict();
        for ncol in range(0, data_sheet.ncols):
            cell = data_sheet.cell(0, ncol);
            column_header = self._get_cell_raw(cell);
            column_headers[column_header] = ncol;

        sheet_desc = _SheetDesc(sheet_name, table_name);
        for row_idx in range(1, meta_sheet.nrows):
            cell = meta_sheet.cell(row_idx, column_idx);
            if cell.ctype != xlrd.XL_CELL_TEXT or cell.value == u"":
                continue;
            text_split = cell.value.split("=");
            column_name = text_split[0];
            field_name = text_split[1];
            if column_name not in column_headers:
                raise Exception("Meta data error, column(%s) not exist in sheet %s" % (column_name, sheet_name));
            sheet_desc.map(column_name, field_name, column_headers[column_name]);
        #不能所有的列都是索引列
        if len(sheet_desc.keys) > 0 and len(sheet_desc.keys) == len(sheet_desc.columns):
            raise Exception("Meta data error, too many keys, sheet: %s" % sheet_name);
        self._meta_tables.append(sheet_desc);
        return True;

    def _load_meta_header(self):
        for sheet_name in self._sheet_names:
            data_sheet = self._workbook.sheet_by_name(sheet_name);
            sheet_desc = _SheetDesc(sheet_name, sheet_name);
            for column_idx in range(0, data_sheet.ncols):
                cell = data_sheet.cell(0, column_idx);
                column_header = self._get_cell_raw(cell);
                if column_header == u"":
                    continue;
                sheet_desc.map(column_header, column_header, column_idx);
            #不能所有的列都是索引
            if len(sheet_desc.keys) == len(sheet_desc.columns):
                raise Exception("Meta data error, too many keys for columns, sheet: %s" % sheet_name);
            self._meta_tables.append(sheet_desc);

    def _convert_sheet(self, sheet_desc):
        if sheet_desc.has_key:
            self._gen_table_code(sheet_desc);
            return;
        self._gen_array_code(sheet_desc);

    #该函数尽可能返回xls看上去的字面值
    def _get_cell_raw(self, cell):
        if cell.ctype == xlrd.XL_CELL_TEXT:
            return cell.value;
        if cell.ctype == xlrd.XL_CELL_NUMBER:
            return (number2string % cell.value).rstrip('0').rstrip('.');
        if cell.ctype == xlrd.XL_CELL_DATE:
            dt = xlrd.xldate.xldate_as_datetime(cell.value, self._workbook.datemode);
            return u"%s" % dt;
        if cell.ctype == xlrd.XL_CELL_BOOLEAN:
            return  u"true" if cell.value else u"false";
        return u"";

    def _get_cell_string(self, cell):
        cell_text = "";
        if cell.ctype == xlrd.XL_CELL_TEXT:
            cell_text = cell.value;
        if cell.ctype == xlrd.XL_CELL_NUMBER:
            cell_text = (number2string % cell.value).rstrip('0').rstrip('.');
        if cell.ctype == xlrd.XL_CELL_DATE:
            dt = xlrd.xldate.xldate_as_datetime(cell.value, self._workbook.datemode);
            cell_text = u"%s" % dt;
        if cell.ctype == xlrd.XL_CELL_BOOLEAN:
            cell_text =  u"true" if cell.value else u"false";
        return u'"%s"' % cell_text;

    def _get_cell_number(self, cell):
        if cell.ctype == xlrd.XL_CELL_TEXT:
            #这里认为用户填的是一个数字,可能是整数,也可能是小数,也可能是十六进制...
            return cell.value;
        if cell.ctype == xlrd.XL_CELL_NUMBER:
            return (number2string % cell.value).rstrip('0').rstrip('.');
        if cell.ctype == xlrd.XL_CELL_DATE:
            dt = xlrd.xldate.xldate_as_datetime(cell.value, self._workbook.datemode);
            return u"%d" % time.mktime(dt.timetuple());
        if cell.ctype == xlrd.XL_CELL_BOOLEAN:
            return  u"1" if cell.value else u"0";
        return u"0";

    def _get_cell_bool(self, cell):
        text = self._get_cell_raw(cell);
        text = text.lower();
        if text in [u"", u"nil", u"0", u"false", u"no", u"none", u"否", u"无"]:
            return u"false";
        return u"true";

    def _gen_array_code(self, sheet_desc):
        self._lines.append(u"--%s: %s\n" % (self._xls_filename, sheet_desc.sheet_name));
        self._lines.append(u"%s%s =\n" % (self._scope == u"local" and u"local " or self._scope == u"global" and u"_G." or u"", sheet_desc.table_name));
        self._lines.append(u"{\n");
        sheet = self._workbook.sheet_by_name(sheet_desc.sheet_name);
        for row in sheet.get_rows():
            line_code = u"    " if len(row) <= 1 else u"    {";
            cell_idx = 1;
            for cell in row:
                if cell_idx != 1:
                    line_code += u", ";
                line_code += self._get_cell_raw(cell);
                cell_idx = cell_idx + 1;
            line_code += u",\n" if len(row) <= 1 else u"},\n";
            self._lines.append(line_code);
        self._lines.append(u"};\n");
        self._lines.append(u"\n");

    def _gen_table_code(self, sheet_desc):
        sheet = self._workbook.sheet_by_name(sheet_desc.sheet_name);
        root = list();
        #生成层级数据结构
        for row_idx in range(1, sheet.nrows):
            row_content = dict();
            for column_desc in sheet_desc.columns:
                row_content[column_desc.field_name] = self._get_cell_text(sheet, row_idx, column_desc);
            node = root;
            for key_idx in range(0, len(sheet_desc.keys)):
                key_desc = sheet_desc.keys[key_idx];
                field_value = row_content[key_desc.field_name];
                child = next((kv["v"] for kv in node if kv["k"] == field_value), None);
                if child == None:
                    #这里用了list,而不是dict,是为了保持最终生成的行顺序尽可能跟填表顺序一致
                    child = list();
                    comment = key_desc.column_name;
                    node.append({"k":field_value, "v":child, "c":comment});
                node = child;
            for column_desc in sheet_desc.columns:
                if not column_desc.is_key:
                    field_name = column_desc.field_name;
                    field_value = row_content[field_name];
                    comment = column_desc.column_name;
                    node.append({"k":field_name, "v":field_value, "c":comment});

        comment = u"%s: %s" % (self._xls_filename, sheet_desc.sheet_name);
        table_var = u"%s%s" % (self._scope == u"local" and u"local " or self._scope == "global" and "_G." or u"", sheet_desc.table_name);
        self._gen_tree_code(sheet_desc, root, 0, table_var, comment);
        self._lines.append(u"\n");

    def _gen_tree_code(self, sheet_desc, node, step, key_name, comment):
        if comment != None:
            self._lines.append(self._indent * step + u"--" + comment + u"\n");

        if step >= len(sheet_desc.keys):
            if len(node) == 1:
                child = node[0];
                line = self._indent * step + key_name + u" = " + child["v"];
                if comment != None:
                    line += u", --%s\n" % child["c"];
                else:
                    line += u",\n";
                self._lines.append(line);
                return;
            line = self._indent * step + key_name + u" = {";
            first_item = True;
            for kv in node:
                lua_name = kv["k"];
                if not first_item:
                    line += u", ";
                line += u"%s=%s" % (lua_name, kv["v"]);
                first_item = False;
            line += u"},\n"
            self._lines.append(line);
            return;

        self._lines.append(self._indent * step + key_name + u" =\n");
        self._lines.append(self._indent * step + u"{\n");
        firstNode = True;
        for kv in node:
            comment = kv["c"] if firstNode else None;
            self._gen_tree_code(sheet_desc, kv["v"], step + 1, u"[%s]" % kv["k"], comment);
            firstNode = False;
        self._lines.append(self._indent * step + u"}" + (u";" if step == 0 else u",") + u"\n");

    def _get_cell_text(self, sheet, row_idx, column_desc):
        cell = sheet.cell(row_idx, column_desc.column_idx);
        if column_desc.map_type == "number":
            return self._get_cell_number(cell);
        if column_desc.map_type == "bool":
            return self._get_cell_bool(cell);
        if column_desc.map_type == "string":
            return self._get_cell_string(cell);
        text = self._get_cell_raw(cell);
        return text if text != "" else "nil";

if __name__ == "__main__":
    parser = argparse.ArgumentParser("excel to lua convertor");
    parser.add_argument("-s", "--scope", dest="scope", help="table scope,local,global", choices=["local", "global", "default"]);
    parser.add_argument("-i", "--indent", dest="indent", help="indent size, 0 for tab, default 4 (spaces)", type=int, default=4, choices=[0, 2, 4, 8]);
    parser.add_argument("-m", "--meta", dest="meta", help="meta sheet name, default 'xls2lua'", default="xls2lua");
    parser.add_argument("-o", "--output", dest="output", help="output file", default="output.lua");
    parser.add_argument("-f", "--force", dest="force", action="store_true", help="force convert");
    parser.add_argument('inputs', nargs='+', help="input excel files");
    args = parser.parse_args();
    converter = Converter(args.scope, args.indent, args.meta);
    if args.force or any(converter.compare_time(filename, args.output) for filename in args.inputs):
        for filename in args.inputs:
            converter.convert(filename);
        converter.save(args.output);

