--xls_sha1=fc088d6f01e52cdcaa9f30154c5e75f009b08bd6
--带列标题示例@test1.xlsx
sheet =
{
    --大段位
    [1] =
    {
        --小段位
        [1] = {elo=0, fight_count=3, protect=true, name="青铜1"},
        [2] = {elo=100, fight_count=3, protect=false, name="青铜2"},
        [3] = {elo=200, fight_count=3, protect=false, name="青铜3"},
    },
    [2] =
    {
        --小段位
        [1] = {elo=300, fight_count=3, protect=true, name="白银1"},
        [2] = {elo=400, fight_count=3, protect=false, name="白银2"},
        [3] = {elo=500, fight_count=3, protect=false, name="白银3"},
    },
    [3] =
    {
        --小段位
        [1] = {elo=600, fight_count=5, protect=true, name="黄金1"},
        [2] = {elo=700, fight_count=5, protect=false, name="黄金2"},
        [3] = {elo=800, fight_count=5, protect=false, name="黄金3"},
    },
};

-- insert some code --
for key, node in pairs(sheet) do
    print(key);
end
