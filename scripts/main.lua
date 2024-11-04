require("config")
require("parse_excel")
require("excel_to_protobuf")

print("hello xlnt lua.")

local configs = parse_args()

for k,v in pairs(configs) do
    -- 
    if "help"==k then
        print_help()
    elseif "excel"==k then
        local parse_excel = ParseExcel(v)
        local proto = ToProtobuf(parse_excel)
        local proto_data_table = ToLuaTable(parse_excel)
        for key, value in pairs(proto_data_table) do
            print("protobuf_encode",key)
            protobuf_encode(proto,key,value)
        end
        
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

