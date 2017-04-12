--xls_sha1=e527908506b8ad732225775eaf4ece9cb602a8b0
--2017-04-12 15:57:05, https://github.com/trumanzhao/xls2lua
--双列示例@test1.xlsx
sheet =
{
    --名字
    ["张三"] = 10 --等级,
    ["李四"] = 20,
    ["王麻子"] = 30,
};

-- insert some code --
for key, node in pairs(sheet) do
    print(key);
end
