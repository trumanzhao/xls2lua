--xls_sha1=e527908506b8ad732225775eaf4ece9cb602a8b0
--数组示例@test1.xlsx
sheet =
{
    {0, 1, 2, 3, 4, 5},
    {1, 1, 2, 3, 4, 5},
    {2, 2, 4, 6, 8, 10},
    {3, 3, 6, 9, 12, 15},
    {4, 4, 8, 12, 16, 20},
    {5, 5, 10, 15, 20, 25},
};

-- insert some code --
for key, node in pairs(sheet) do
    print(key);
end
