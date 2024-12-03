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
    local proto,proto_package,package_name = parse_excel.ToProtobuf()
    local proto_parse = "syntax = \"proto3\";\n\n"..proto
     --保存proto文件
     local proto_path = out_dir.."/"..excel_name..".proto"
     local proto_content = "syntax = \"proto3\";\n\n"..proto_package..proto
     local proto_file = io.open(proto_path,"w")
     if proto_file ~= nil then
         proto_file:write(proto_content)
         proto_file:close()
         log.info("lua proto write success. -> ",proto_path)
     end
    -- protobuf 
    local proto_data_table = parse_excel.ToJson()
    for key, value in pairs(proto_data_table) do
        local config_name = value.name
        if target == nil or target == config_name then
            local proto_type = key
            if package_name ~= nil and #package_name > 0 then
                proto_type = package_name.."."..proto_type
            end
            local textproto = string.gsub(value.textproto,"\"","\\\"")
            local binary_path = out_dir.."/"..config_name..".bytes"
            local protoc_cmd_format = "echo \"%s\" | protoc --encode=%s %s > %s"
            local protoc_cmd = string.format(protoc_cmd_format,textproto,proto_type,proto_path,binary_path)
            log.info(protoc_cmd)
            local pc = io.popen(protoc_cmd)
            if pc ~= nil then
                -- 读取命令的输出结果
                local pc_result = pc:read("*a")
                pc:close()
                log.info(pc_result)
            end
        end
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



