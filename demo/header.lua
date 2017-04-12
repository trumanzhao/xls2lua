--xls_sha1=e527908506b8ad732225775eaf4ece9cb602a8b0
--2017-04-12 15:57:05, https://github.com/trumanzhao/xls2lua
--带列标题示例@test1.xlsx
sheet =
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

-- insert some code --
for key, node in pairs(sheet) do
    print(key);
end
