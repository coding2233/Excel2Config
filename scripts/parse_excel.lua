-- require ("parse_excel_template")

local parse_excel = {}
local this = parse_excel

local function ParseMessage(vSheet,kRow,kCell)

    --检查
    if vSheet[kRow] == nil or vSheet[kRow+1] == nil or vSheet[kRow+2] == nil or vSheet[kRow+3] == nil then
        print("excel message template cannot be parsed normally")
        return kRow,kCell
    end

    local type = vSheet[kRow][kCell+1]
    if type == nil or string.len(type) == 0 or #type==0 then
        return kRow,kCell
    end

    --message的变量
    local var_list = {}
    local var_index = 1
    while true do
        local var_type = vSheet[kRow+1][kCell+var_index]
        if var_type == nil or #var_type == 0 or string.len(var_type) == 0 then
            -- print("while var_list break")
            break
        end
        
        local var_name = vSheet[kRow+2][kCell+var_index]
        local var_desc = vSheet[kRow+3][kCell+var_index]
        local var_data = {type = var_type, var = var_name, desc = var_desc}
        table.insert(var_list,var_data)
        var_index = var_index+1
    end

   
    -- 消息模板
    local message_template = {type = type,row=kRow,cloumn = kCell, var_list = var_list}
    table.insert(parse_excel.message_template_list,message_template)

    -- 配置入口
    if vSheet[kRow-1] ~= nil and vSheet[kRow-1][kCell] ~= nil and #"config" == vSheet[kRow-1][kCell] then
        local config_value = vSheet[kRow-1][kCell+1]
        if config_value ~= nil then
            parse_excel.excel_config[type] = config_value
        end
    end

    return kRow,kCell
end

local function ParseEnum(vSheet,kRow,kCell)
    local enum_type = vSheet[kRow][kCell+1]
    local enmu_desc = vSheet[kRow+1][kCell+1]
    local enum_value_list = {}
    local enum_value_index = 3
    while true do
        if vSheet[kRow+enum_value_index] == nil then
            break
        end
        local enmu_var = vSheet[kRow+enum_value_index][kCell+1]
        if enmu_var==nil or #enmu_var == 0 or string.len(enmu_var) == 0 then
            break
        end
        local enmu_var_desc = vSheet[kRow+enum_value_index][kCell+2]
        local enmu_var_value = vSheet[kRow+enum_value_index][kCell+3]
        enum_value_index = enum_value_index + 1

        local enum_value = {var =enmu_var,desc =enmu_var_desc,value = enmu_var_value }
        table.insert(enum_value_list,enum_value)
    end

    local enum_type = {type=enum_type,desc=enmu_desc,value_list=enum_value_list}
    table.insert(parse_excel.enmu_list,enum_type)
    return kRow+enum_value_index,1
end

local function ParsePackage(vSheet,kRow,kCell)
    parse_excel.package = vSheet[kRow][kCell+1]
    -- print("ParsePackage",parse_excel.package)

    return kRow,kCell+1
end

local function ParseExcelCell(vSheet)
    for row = 1, #vSheet do
        local row_data = vSheet[row]
        for cloumn = 1, #row_data do
            local next_row = row
            local next_cloumn = cloumn
            local cell = row_data[cloumn]
            if "#message" == cell then
                next_row,next_cloumn = ParseMessage(vSheet,row,cloumn)
            -- 拼写错误
            elseif "#enum" == cell or "#enmu" == cell then
                next_row,next_cloumn = ParseEnum(vSheet,row,cloumn)
            elseif "#package" == cell then
                next_row,next_cloumn = ParsePackage(vSheet,row,cloumn)
            end
            -- 更新读取excel的行列索引
            cloumn = next_cloumn
            if row ~= next_row then
                row = next_row
                break
            end
        end
    end
end

local function ParseExcelFile(excel_path)
    local excel = read_excel(excel_path)
    -- sheet
    for kSheet,vSheet in pairs(excel) do
        -- print(kSheet,vSheet)
        ParseExcelCell(vSheet)
    end
end

function ParseExcel(path)
    -- print("ParseExcel",path)

    parse_excel = {}
    parse_excel.package = nil
    parse_excel.message_template_list = {}
    parse_excel.enmu_list = {}
    parse_excel.excel_config = {}

    local files = get_files(path)
    for i = 1,#files do
        ParseExcelFile(files[i])
    end

    return parse_excel
end

local function MessageToProtobuf(message_template)
    local string_builder = {}
    -- local message_template = {type = type,row=kRow,cloumn = kCell, var_list = var_list}
    -- local var_data = {type = var_type, var = var_name, desc = var_desc}
    local type = message_template.type
    table.insert(string_builder,string.format("message %s\n{\n",type))
    for i=1, #message_template.var_list do
        local message_var = message_template.var_list[i]
        if message_var.desc ~= nil and #message_var.desc > 0 then
            table.insert(string_builder, string.format("\t//%s\n",message_var.desc))
        end
        table.insert(string_builder, string.format("\t%s=%s;\n",message_var.var,tostring(i)))
    end
    table.insert(string_builder, "}\n\n")

    return table.concat(string_builder)
end

local function EnumToProtobuf(enum_template)
    --local enum_type = {type=enum_type,desc=enmu_desc,value_list=enum_value_list}
    local string_builder = {}
    local type = enum_template.type
    local desc = enum_template.desc
    local value_list = enum_template.value_list
    --local enum_value = {var =enmu_var,desc =enmu_var_desc,value = enmu_var_value }

    if desc ~= nil and #desc > 0 then
        table.insert(string_builder, string.format("//%s\n",desc))
    end
    table.insert(string_builder, string.format("enum %s\n{\n",type))
    for i = 1,#enum_template.value_list do
        local enum_value = enum_template.value_list[i]
        if enum_value.desc ~= nil and #enum_value.desc > 0 then
            table.insert(string_builder, string.format("\t//%s\n",enum_value.desc))
        end
        table.insert(string_builder, string.format("\t%s=%s;\n",enum_value.var,enum_value.value))
    end
    table.insert(string_builder, "}\n\n")

    return table.concat(string_builder)
end

function ToProtobuf(excel_template)
    -- print("ToProtobuf")
    local string_builder = {}
    table.insert(string_builder,"syntax = \"proto3\";\n\n")
    if excel_template.package ~= nil and string.len(excel_template.package) > 0 then
        table.insert(string_builder,string.format("package %s;\n\n",excel_template.package))
    end

    -- EnumToProtobuf
    for i=1,#parse_excel.enmu_list do
        local enmu_string = EnumToProtobuf(parse_excel.enmu_list[i])
        if enmu_string~=nil and string.len(enmu_string) > 0 then
            table.insert(string_builder,enmu_string)
        end
    end

    -- MessageToProtobuf
    for i=1,#parse_excel.message_template_list do
        local message_string = MessageToProtobuf(parse_excel.message_template_list[i])
        if message_string~=nil and string.len(message_string) > 0 then
            table.insert(string_builder,message_string)
        end
    end

    local protobuf_string = table.concat(string_builder)
    print(protobuf_string)

end


function ToLuaTable(excel_template)
    
end