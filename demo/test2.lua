--xls_sha1=e50b0c13db61209670d7dd19bea8a304afeb3549
--2017-04-12 15:57:05, https://github.com/trumanzhao/xls2lua
--test2@test2.xlsx
local sheet =
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
return sheet;
