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
            local value_name = value.name
            local value_data = value.data
            print("protobuf_encode",key)
            local data_path = get_exe_dir().."/package/data_temp.lua"
            local data_file = io.open(data_path,"w")
            data_file:write(value_data)
            data_file:close()

            local data_table = require("data_temp")
            local bytes = ProtobufExcelEncode(proto,key,data_table)
            -- PBTest()

            local binary_path = get_exe_dir().."/"..value_name..".bytes"
            local binary_file = io.open(binary_path,"w")
            binary_file:write(bytes)
            binary_file:close()
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

