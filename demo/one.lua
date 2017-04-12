--xls_sha1=e527908506b8ad732225775eaf4ece9cb602a8b0
--2017-04-12 15:57:05, https://github.com/trumanzhao/xls2lua
--单列示例@test1.xlsx
sheet =
{
    "无类型指示的字符串必须加引号",
    "单列Sheet是无法做映射的,当做数组",
    "张三",
    "李四",
    "王麻子",
    "当然也可以是下面这样的数字",
    12345,
};

-- insert some code --
for key, node in pairs(sheet) do
    print(key);
end
