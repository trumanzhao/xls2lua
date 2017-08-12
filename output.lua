--2017-08-13 01:23:07, https://github.com/trumanzhao/xls2lua

--带列标题示例@test1.xlsx
table_with_header =
{
    --大段位
    [1] =
    {
        --小段位
        [1] = {elo=0, name="青铜1", fight_count=3, reward=nil, protect=true},
        [2] = {elo=100, name="青铜2", fight_count=3, reward={102, 1}, protect=false},
        [3] = {elo=200, name="青铜3", fight_count=3, reward={103, 2}, protect=false},
    },
    [2] =
    {
        --小段位
        [1] = {elo=300, name="白银1", fight_count=3, reward={104, 3}, protect=true},
        [2] = {elo=400, name="白银2", fight_count=3, reward={105, 4}, protect=false},
        [3] = {elo=500, name="白银3", fight_count=3, reward={106, 5}, protect=false},
    },
    [3] =
    {
        --小段位
        [1] = {elo=600, name="黄金1", fight_count=5, reward={107, 6}, protect=true},
        [2] = {elo=700, name="黄金2", fight_count=5, reward={108, 7}, protect=false},
        [3] = {elo=800, name="黄金3", fight_count=5, reward={109, 8}, protect=false},
    },
};

--单列示例@test1.xlsx
sigle_column =
{
    "无类型指示的字符串必须加引号",
    "单列Sheet是无法做映射的,当做数组",
    "张三",
    "李四",
    "王麻子",
    "当然也可以是下面这样的数字",
    12345,
};

--双列示例@test1.xlsx
dictionary =
{
    --名字
    ["张三"] = 10, --等级
    ["李四"] = 20,
    ["王麻子"] = 30,
};

--数组示例@test1.xlsx
array_data =
{
    {0, 1, 2, 3, 4, 5},
    {1, 1, 2, 3, 4, 5},
    {2, 2, 4, 6, 8, 10},
    {3, 3, 6, 9, 12, 15},
    {4, 4, 8, 12, 16, 20},
    {5, 5, 10, 15, 20, 25},
};

--test2@test2.xlsx
test2 =
{
    --*dan#
    [1] =
    {
        --*step#
        [1] = {min_elo=0, fight=3, name="青铜1", protect=true},
        [2] = {min_elo=100, fight=3, name="青铜2", protect=false},
        [3] = {min_elo=200, fight=3, name="青铜3", protect=false},
    },
    [2] =
    {
        --*step#
        [1] = {min_elo=300, fight=3, name="白银1", protect=true},
        [2] = {min_elo=400, fight=3, name="白银2", protect=false},
        [3] = {min_elo=500, fight=3, name="白银3", protect=false},
    },
    [3] =
    {
        --*step#
        [1] = {min_elo=600, fight=5, name="黄金1", protect=true},
        [2] = {min_elo=700, fight=5, name="黄金2", protect=false},
        [3] = {min_elo=800, fight=5, name="黄金3", protect=false},
    },
};

