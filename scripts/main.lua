require("config")

print("hello xlnt lua.")

local configs = parse_args()

for k,v in pairs(configs) do
    print(k,v)
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

