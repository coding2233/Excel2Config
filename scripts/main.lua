require("config")
require("lfs")
local parse_excel = require("parse_excel")

log = require "log"
log.level = "info"


--执行目录
local exec_dir = get_exe_dir()

local function parse_excel_to_protobuf(excel_file,excel_name,out_dir,target)
    log.info(string.format("parse_excel_to_protobuf\n excel: %s\n excel_name: %s\n out_dir: %s",excel_file,excel_name,out_dir))
    --加载excel
    parse_excel.Load(excel_file)
    -- proto
    local proto,proto_package = parse_excel.ToProtobuf()
    local proto_parse = "syntax = \"proto3\";\n\n"..proto
    -- protobuf 
    local proto_data_table = parse_excel.ToLuaTable()
    for key, value in pairs(proto_data_table) do
        local value_name = value.name
        if target == nil or target == value_name  then
            local value_data = value.data
            -- log.debug("protobuf_encode",key,value_name,value_data)
            -- 将protobuf lua table写入一个临时的lua文件
            local data_path = exec_dir.."/package/data_temp.lua"
            local data_file = io.open(data_path,"w")
            if data_file ~= nil then
                data_file:write(value_data)
                data_file:close()
                log.debug("write succee->",key,value_name,data_path)
            end
            -- 读取protobuf的lua table
            local data_table = require("data_temp")
            log.info("lua table write success. -> ",data_path,data_table)
            local excel_pb = require("excel_to_protobuf")
            local bytes = excel_pb.ProtobufExcelEncode(proto_parse,key,data_table)
            -- PBTest()

            -- 生成配置的二进制文件
            local binary_path = out_dir.."/"..value_name..".bytes"
            local binary_file = io.open(binary_path,"w")
            if binary_file ~= nil then
                binary_file:write(bytes)
                binary_file:close()
                log.info("lua protobuf binary write success. -> ",binary_path)
            end
        end
    end
    --保存proto文件
    local proto_path = out_dir.."/"..excel_name..".proto"
    local proto_content = "syntax = \"proto3\";\n\n"..proto_package..proto
    local proto_file = io.open(proto_path,"w")
    if proto_file ~= nil then
        proto_file:write(proto_content)
        proto_file:close()
        log.info("lua proto write success. -> ",proto_path)
    end
end

local function find_excel(excel_dir,out_dir,target)
    if excel_dir == nil then
        log.error("excel_dir is nil.")
        return
    end

    if out_dir == nil then
        out_dir = excel_dir
    end

    for excel_file in lfs.dir(excel_dir) do
        log.debug(excel_file)
        local excel_name = excel_file
        local a,b = string.find(excel_file,".xlsx")
        if a ~= nil then
            excel_name = string.gsub(excel_file,".xlsx","")
            log.debug(excel_file,#excel_file,a,b,excel_name)
            parse_excel_to_protobuf(excel_dir.."/"..excel_file,excel_name,out_dir,target)
        end
    end
end

local function parse_args()
    local excel2config_args = {}
    local configs = ParseArgs()
    for k,v in pairs(configs) do
        -- 
        if "help"==k then
            PrintHelp()
            return nil
        elseif "version"==k then
            print("version",VERSION)
            return
        elseif "debug"==k then
            log.level = "trace"
        elseif "excel"==k then
            excel2config_args.excel_dir = v
        elseif "out"==k then
            excel2config_args.out_dir = v
        elseif "target"==k then
            excel2config_args.target = v
        end
    end
    return excel2config_args
end


local function run()
    local excel2config_args = parse_args()
    if excel2config_args ~= nil then
        find_excel(excel2config_args.excel_dir,excel2config_args.out_dir, excel2config_args.target)
    end
end

-- 执行
local status,error = pcall(run())

log.info("run",status,error)



