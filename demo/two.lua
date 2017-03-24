--xls_sha1=fc088d6f01e52cdcaa9f30154c5e75f009b08bd6
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
