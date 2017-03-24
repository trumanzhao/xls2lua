#!/usr/bin/python
# -*- coding: utf-8 -*-
#repository: https://github.com/trumanzhao/xls2lua
#trumanzhao, 2017/03/24, trumanzhao@foxmail.com

import os, time, hashlib, codecs, xlrd

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

class _ColumnDesc(object):
    """列描述"""
    def __init__(self, column_name, lua_name, column_idx):
        first_char = lua_name[0];
        last_char = lua_name[-1];
        map_table = {u"?":"bool", u"#":"number", u"$":"string"};
        lua_name = lua_name if first_char != u"*" else lua_name[1:];
        lua_name = lua_name if last_char not in map_table else lua_name[:-1];
        self.column_name = column_name;
        self.column_idx = column_idx;
        self.is_key = (first_char == u"*");
        self.lua_name = lua_name;
        self.map_type = map_table[last_char] if last_char in map_table else "raw";

class _SheetDesc(object):
    """sheet描述"""
    def __init__(self, sheet_name, lua_name):
        self.sheet_name = sheet_name;
        self.lua_name = lua_name;
        self.columns = list();
        self.maps = dict();
        self.keys = list();
        self.has_key = False;

    def map(self, column_name, lua_name, column_idx):
        desc = _ColumnDesc(column_name, lua_name, column_idx);
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
    _tab_step = u" " * 4;
    _out_dir = u"";
    _check_hash = True;
    _code_writer = None;
    _hash_tag = "--xls_sha1=";
    _local_sheet = False;
    _return_sheet = False;

    #tab_step: 生成lua代码的tab缩进,默认为四个空格(字符串),可以设置为任意个空格或'\t'的字符串
    #out_dir: 输出路径,默认当前目录
    #check_hash: 是否检查文件哈希
    #local_sheet: 是否将sheet对应的table名作为local变量,默认不加
    #return_sheet: 是否在代码最后return sheet,默认不加;如果需要在末尾插入代码的话,建议不要打开这项.
    def __init__(self, **kwargs):
        if "tab_step" in kwargs:
            self._tab_step = _unicode_anyway(kwargs["tab_step"]);
        if "out_dir" in kwargs:
            self._out_dir = _unicode_anyway(kwargs["out_dir"]);
        if "check_hash" in kwargs:
            self._check_hash = kwargs["check_hash"];
        if "local_sheet" in kwargs:
            self._local_sheet = kwargs["local_sheet"];
        if "return_sheet" in kwargs:
            self._return_sheet = kwargs["return_sheet"];

    #code_writer: lua代码的自定义写入器,一般是用来对代码做写入前的修改
    #def code_writer(sheet_name, lua_path, code): return modified(code);
    #其中, sheet_name为excel的sheet name, lua_path为写入的lua文件名, code是Unicode字符串
    def convert(self, xls_filename, code_writer=None):
        xls_filename = _unicode_anyway(xls_filename);
        try:
            self._workbook = xlrd.open_workbook(xls_filename);
            self._xls_filetime = os.path.getmtime(xls_filename);
            self._xls_filename = xls_filename;
            self._xls_hash = hashlib.sha1(open(xls_filename, 'rb').read()).hexdigest()
        except:
            raise Exception("Failed to load workbook: %s" % xls_filename);

        self._code_writer = code_writer;
        self._sheet_names = self._workbook.sheet_names();
        self._meta_tables = list();
        if "xls2lua" in self._sheet_names:
            self._load_meta_sheet();
        else:
            self._load_meta_header();

        for sheet_desc in self._meta_tables:
            self._convert_sheet(sheet_desc);

    #meta_tables: list of _SheetDesc
    #meta_tables之所以是一个list而不是dict,是因为允许对同一个sheet做多个映射转换
    def _load_meta_sheet(self):
        meta_sheet = self._workbook.sheet_by_name("xls2lua");
        for column_idx in range(0, meta_sheet.ncols):
            self._load_meta_column(meta_sheet, column_idx);

    #meta_sheet中,每两列定义了一个sheet的映射
    #本函数将这两列数据load为一个meta_table:
    def _load_meta_column(self, meta_sheet, column_idx):
        text = meta_sheet.cell(0, column_idx).value;
        text_split = text.split("=");
        lua_name = text_split[0];
        sheet_name = text_split[1];
        if sheet_name not in self._sheet_names:
            raise Exception("Meta error, sheet not exist: %s" % sheet_name);

        data_sheet = self._workbook.sheet_by_name(sheet_name);
        column_headers = dict();
        for ncol in range(0, data_sheet.ncols):
            cell = data_sheet.cell(0, ncol);
            column_header = self._get_cell_raw(cell);
            column_headers[column_header] = ncol;

        sheet_desc = _SheetDesc(sheet_name, lua_name);
        for row_idx in range(1, meta_sheet.nrows):
            cell = meta_sheet.cell(row_idx, column_idx);
            if cell.ctype != xlrd.XL_CELL_TEXT or cell.value == u"":
                continue;
            text_split = cell.value.split("=");
            lua_name = text_split[0];
            column_name = text_split[1];
            if column_name not in column_headers:
                raise Exception("Meta data error, column(%s) not exist in sheet %s" % (column_name, sheet_name));
            sheet_desc.map(column_name, lua_name, column_headers[column_name]);
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

    def _get_lua_path(self, lua_name):
        lua_path = lua_name;
        ext = os.path.splitext(lua_path)[1].lower();
        if ext != ".lua":
            lua_path += ".lua";
        lua_path = os.path.join(self._out_dir, lua_path);
        lua_path = os.path.normpath(lua_path);
        return lua_path;

    def _convert_sheet(self, sheet_desc):
        lua_path = self._get_lua_path(sheet_desc.lua_name);
        stored_hash = self._get_stored_hash(lua_path);
        if self._check_hash and stored_hash == self._xls_hash:
            #文件没有变化,无需转换
            return;
        if not sheet_desc.has_key:
            self._gen_array_code(sheet_desc);
            return;
        self._gen_table_code(sheet_desc);

    def _get_stored_hash(self, lua_path):
        try:
            line = codecs.open(lua_path, 'r', "utf-8").readline().strip();
            if line.startswith(self._hash_tag):
                return line[len(self._hash_tag):];
            return "";
        except:
            return "";

    #尽可能把cell按其在excel中看起来的样子读成字符串
    def _get_cell_raw(self, cell):
        if cell.ctype == xlrd.XL_CELL_TEXT:
            return cell.value;
        if cell.ctype == xlrd.XL_CELL_NUMBER:
            number = int(cell.value);
            return u"%d" % number if number == cell.value else u"%g" % cell.value;
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
            number = int(cell.value);
            cell_text = u"%d" % number if number == cell.value else u"%g" % cell.value;
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
            number = int(cell.value);
            return u"%d" % number if number == cell.value else u"%g" % cell.value;
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
        lines = list();
        lines.append(u"%s%s\n" % (self._hash_tag, self._xls_hash));
        lines.append(u"--%s@%s\n" % (sheet_desc.sheet_name, self._xls_filename));
        lines.append(u"%ssheet =\n" % (u"local " if self._local_sheet else u""));
        lines.append(u"{\n");
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
            lines.append(line_code);
        lines.append(u"};\n");
        if self._return_sheet:
            lines.append(u"return sheet;\n");
        self._write_lua(sheet_desc, ''.join(lines));

    def _gen_table_code(self, sheet_desc):
        sheet = self._workbook.sheet_by_name(sheet_desc.sheet_name);
        root = list();
        #生成层级数据结构
        for row_idx in range(1, sheet.nrows):
            row_content = dict();
            for column_desc in sheet_desc.columns:
                row_content[column_desc.lua_name] = self._get_cell_text(sheet, row_idx, column_desc);
            node = root;
            for key_idx in range(0, len(sheet_desc.keys)):
                key_desc = sheet_desc.keys[key_idx];
                lua_value = row_content[key_desc.lua_name];
                child = next((kv["v"] for kv in node if kv["k"] == lua_value), None);
                if child == None:
                    #这里用了list,而不是dict,是为了保持最终生成的行顺序尽可能跟填表顺序一致
                    child = list();
                    comment = key_desc.column_name;
                    node.append({"k":lua_value, "v":child, "c":comment});
                node = child;
            for column_desc in sheet_desc.columns:
                if not column_desc.is_key:
                    lua_name = column_desc.lua_name;
                    lua_value = row_content[lua_name];
                    comment = column_desc.column_name;
                    node.append({"k":lua_name, "v":lua_value, "c":comment});

        lines = list();
        lines.append(u"%s%s\n" % (self._hash_tag, self._xls_hash));
        comment = u"%s@%s" % (sheet_desc.sheet_name, self._xls_filename);
        table_var = u"%ssheet" % (u"local " if self._local_sheet else u"");
        self._gen_tree_code(lines, sheet_desc, root, 0, table_var, comment);
        if self._return_sheet:
            lines.append(u"return sheet;\n");
        self._write_lua(sheet_desc, ''.join(lines));

    def _write_lua(self, sheet_desc, code):
        sheet_name = sheet_desc.sheet_name;
        lua_path = self._get_lua_path(sheet_desc.lua_name);
        #try:
        lua_dir = os.path.split(lua_path)[0];
        if lua_dir != "" and not os.path.exists(lua_dir):
            os.makedirs(lua_dir)
        if self._code_writer != None:
            self._code_writer(sheet_name, lua_path, code);
            return;
        open(lua_path, "wb").write(code.encode("utf-8"));
        #except:
        #    raise Exception("Failed to write lua, sheet=%s, lua=%s" % (sheet_name, lua_path));

    def _gen_tree_code(self, lines, sheet_desc, node, step, key_name, comment):
        if comment != None:
            lines.append(self._tab_step * step + u"--" + comment + u"\n");

        if step >= len(sheet_desc.keys):
            if len(node) == 1:
                child = node[0];
                line = self._tab_step * step + key_name + u" = " + child["v"];
                if comment != None:
                    line += " --" + child["c"];
                line += u",\n";
                lines.append(line);
                return;
            line = self._tab_step * step + key_name + u" = {";
            first_item = True;
            for kv in node:
                lua_name = kv["k"];
                if not first_item:
                    line += u", ";
                line += u"%s=%s" % (lua_name, kv["v"]);
                first_item = False;
            line += u"},\n"
            lines.append(line);
            return;

        lines.append(self._tab_step * step + key_name + u" =\n");
        lines.append(self._tab_step * step + u"{\n");
        firstNode = True;
        for kv in node:
            comment = kv["c"] if firstNode else None;
            self._gen_tree_code(lines, sheet_desc, kv["v"], step + 1, u"[%s]" % kv["k"], comment);
            firstNode = False;
        lines.append(self._tab_step * step + u"}" + (u";" if step == 0 else u",") + u"\n");

    def _get_cell_text(self, sheet, row_idx, column_desc):
        cell = sheet.cell(row_idx, column_desc.column_idx);
        if column_desc.map_type == "number":
            return self._get_cell_number(cell);
        if column_desc.map_type == "bool":
            return self._get_cell_bool(cell);
        if column_desc.map_type == "string":
            return self._get_cell_string(cell);
        return self._get_cell_raw(cell);


def _test_writer(sheet_name, lua_path, code):
    code += u'''
-- insert some code --
for key, node in pairs(sheet) do
    print(key);
end
''';
    open(lua_path, "wb").write(code.encode("utf-8"));

if __name__ == "__main__":
    converter = Converter(tab_step=" " * 4, check_hash=False);
    converter.convert('test1.xlsx', _test_writer);

    converter = Converter(tab_step=" " * 4, check_hash=False, local_sheet=True, return_sheet=True);
    converter.convert('test2.xlsx');

