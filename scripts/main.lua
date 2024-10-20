require("config")
require("parse_excel")

print("hello xlnt lua.")

local configs = parse_args()

for k,v in pairs(configs) do
    -- 
    if "help"==k then
        print_help()
    elseif "excel"==k then
        ParseExcel(v)
    end
end

-- function test_require()
--     local pb = require("pb")
--     print(pb)
-- end

-- local status,error = pcall(test_require)
-- print(error)



print("[end] xlnt lua")
-- local test_reult = test_register();
-- print(test_reult)

