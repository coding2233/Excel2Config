-- require ("parse_excel_template")

local parse_excel = {}

local var_base_type_list = {"int32","string","bool","float"}

local function TableConcatEx(table_data)
    local string_builder_result = table.concat(table_data)
    local i,result_j = string.find(string_builder_result,",",-1)
    if result_j == #string_builder_result then
        string_builder_result = string.sub(string_builder_result,1,result_j-1)
    end
    return string_builder_result
end

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
        local var_data = {type = var_type, var = var_name, desc = var_desc,row = kRow+2,cloumn = kCell+var_index}
        table.insert(var_list,var_data)
        var_index = var_index+1
    end

   
    -- 消息模板
    -- print(type,kRow,kCell,#var_list)
    local message_template = {sheet=vSheet,type = type,row=kRow,cloumn = kCell, var_list = var_list}
    table.insert(parse_excel.message_template_list,message_template)

    -- 配置入口
    if vSheet[kRow-1] ~= nil and vSheet[kRow-1][kCell] ~= nil and "#config" == vSheet[kRow-1][kCell] then
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
            elseif "#enum" == cell then
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

function ParseExcel(excel_path)
    -- print("ParseExcel",path)

    parse_excel = {}
    parse_excel.package = nil
    parse_excel.message_template_list = {}
    parse_excel.enmu_list = {}
    parse_excel.excel_config = {}

    ParseExcelFile(excel_path)

    return parse_excel
end

local function MessageToProtobuf(message_template)
    -- print("message_template",message_template,message_template.type,#message_template.var_list)
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
        local var_string = message_var.var
        local var_type_string = message_var.type
        -- map<string,string> 
        -- repeated int32
        local sub_index = string.find(var_type_string,"#")
        -- print("--------",var_type_string,sub_index)
        if sub_index ~= nil and sub_index > 1 then
            var_type_string = string.sub(var_type_string,1,sub_index-1)
        end
        table.insert(string_builder, string.format("\t%s %s = %s;\n",var_type_string,var_string,tostring(i)))
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
    local string_package = nil
    -- table.insert(string_builder,"syntax = \"proto3\";\n\n")
    if excel_template.package ~= nil and string.len(excel_template.package) > 0 then
        -- table.insert(string_builder,string.format("package %s;\n\n",excel_template.package))
        string_package = string.format("package %s;\n\n",excel_template.package)
    end

    -- EnumToProtobuf
    for i=1,#excel_template.enmu_list do
        local enmu_string = EnumToProtobuf(excel_template.enmu_list[i])
        if enmu_string~=nil and string.len(enmu_string) > 0 then
            table.insert(string_builder,enmu_string)
        end
    end

    -- MessageToProtobuf
    for i=1,#excel_template.message_template_list do
        local message_string = MessageToProtobuf(excel_template.message_template_list[i])
        if message_string~=nil and #message_string > 0 then
            -- print(message_string)
            table.insert(string_builder,message_string)
        end
    end

    -- print(#string_builder)
    local protobuf_string = table.concat(string_builder)
    print(protobuf_string)
    return protobuf_string, string_package
end

function GetMessageTemplate(excel_template,message_name)
    for i=1,#excel_template.message_template_list do
        local message_template = excel_template.message_template_list[i]
        if message_name == message_template.type then
            return message_template
        end
    end
    return nil
end



function MessageTypeVarToLua(message_var,excel_template)
    local string_builder = {}
    local type = message_var.type
    local var  = message_var.var
    local row = message_var.row
    local cloumn = message_var.cloumn

    local type_name = type
    local is_list = false
    local i,j = string.find(type,"repeated ")
    if j ~= nil then
        type_name = string.sub(type,j+1)
        is_list = true
    end

    local message_template = GetMessageTemplate(excel_template,type_name)
    -- print("MessageTypeVarToLua",type,type_name,is_list,var,message_template)

    if message_template ~= nil then
        -- for key, value in pairs(message_template) do
        --     print(key,value)
        -- end
        local find_message_template_row = message_template.row
        -- print( message_template.row,message_template.cloumn, message_template.type,message_template.sheet)
        table.insert(string_builder,string.format("%s={",var))
        if is_list then
            local find_row_index = 4;
            while true do
                local find_row_data = message_template.sheet[find_message_template_row+find_row_index]
                -- print(find_message_template_row,find_row_index,find_row_data)
                if find_row_data ~= nil then
                    -- todo 检查这一行都是空数据 100太假了
                    local is_all_nil = true
                    for i=1,100 do
                        if find_row_data[i]~=nil and #find_row_data[i] > 0 then
                            is_all_nil = false
                            break
                        end
                    end

                    -- 空行数据
                    if is_all_nil then
                        break
                    end
                    table.insert(string_builder,"{")
                    -- print(#message_template.var_list)
                    for i=1,#message_template.var_list do
                        local message_var_string = MessageVarToLua(message_template.var_list[i],excel_template,find_row_data)
                        if message_var_string ~= nil and #message_var_string > 0 then
                            table.insert(string_builder,string.format("%s,",message_var_string))
                        end
                    end
                    table.insert(string_builder,"},")
                else
                    -- 空行数据
                    break
                end
                find_row_index=find_row_index+1
            end
            
        else
            local find_row_data = message_template.sheet[find_message_template_row+4]

            for i=1,#message_template.var_list do
                local message_var_string = MessageVarToLua(message_template.var_list[i],excel_template,find_row_data)
                if message_var_string ~= nil and #message_var_string > 0 then
                    table.insert(string_builder,message_var_string)
                end
            end
        end
        table.insert(string_builder,string.format("},",var))
    end

    return TableConcatEx(string_builder)
end



function MessageMapVarToLua(message_var,excel_template,row_data_target)
    local string_builder = {}
    local type = message_var.type
    local var  = message_var.var
    local row = message_var.row
    local cloumn = message_var.cloumn
    
    local var_value = row_data_target[cloumn]
    if var_value == nil then
        var_value = ""
    end

    table.insert(string_builder,"{")
    for k, v in string.gmatch(var_value, "(%w+):(%w+)") do
        table.insert(string_builder,string.format("%s=%s,",k,v))
    end
    table.insert(string_builder,"},")

    return TableConcatEx(string_builder)
end

function MessageBaseVarToLua(message_var,excel_template,row_data_target)
    local string_builder = {}
    local type = message_var.type
    local var  = message_var.var
    local row = message_var.row
    local cloumn = message_var.cloumn
    local type_string = nil
    local is_list = false
    local type_is_string = false
    local type_is_bool = false
    for i=1, #var_base_type_list do
        if type == var_base_type_list[i] then
            type_string = var_base_type_list[i]
            break
        else
            local var_base_list_type = "repeated "..var_base_type_list[i]
            if string.find(type,var_base_list_type) == 1 then
                type_string = var_base_list_type
                is_list = true
                break
            end
        end
    end

    if type_string ~= nil then
        -- print(type_string)
        local find_type_is_string = string.find(type_string,"string")
        type_is_string = find_type_is_string ~= nil and find_type_is_string >0

        local find_type_is_bool = string.find(type_string,"bool")
        type_is_bool = find_type_is_bool ~= nil and find_type_is_bool > 0

        local row_data = row_data_target
        if row_data == nil then
            row_data = excel_template[row+2]
        end
        
        if row_data ~= nil then
            local var_value = row_data[cloumn]
            -- 需要处理一下默认数据
            if var_value== nil or #var==0 then
                if not is_list then
                    if type_is_string then
                        var_value = ""
                    elseif type_is_bool then
                        var_value = false
                    else
                        var_value = 0
                    end
                end
            end
            table.insert(string_builder,string.format("%s=",var))
            if is_list then
                -- todo ..  
                -- e.g.
                -- repeated int32#sep=,
                local find_key = ","
                table.insert(string_builder,"{")
                local read_while= var_value ~= nil and string.len(var_value) > 0
                local read_index = 1
                while read_while do
                    local find_index = string.find(var_value,find_key,read_index)
                    if find_index == nil then
                        find_index = string.len(var_value) + 1 
                        read_while = false
                    end

                    local sub = string.sub(var_value,read_index,find_index-1)

                    local var_vale_format = "%s,"
                    if type_is_string then
                        var_vale_format = "\"%s\","
                    end
                    if type_is_bool then
                        sub = string.lower(sub)
                    end

                    table.insert(string_builder,string.format(var_vale_format,sub))
                    read_index = find_index + 1
                end
                table.insert(string_builder,"},")
            else
                local var_vale_format = "%s,"
                if type_is_string then
                    var_vale_format = "\"%s\","
                end
                if type_is_bool then
                    var_value = string.lower(var_value)
                end
                table.insert(string_builder,string.format(var_vale_format,var_value))
            end
        end
    end

    if #string_builder > 0 then
        return TableConcatEx(string_builder)
    else
        -- 继续处理map类型
        return MessageMapVarToLua(message_var,excel_template,row_data)
    end
end

local function ConfigToLuaTable(message_name,config_name,excel_template)
    local string_builder = {}
    table.insert(string_builder,string.format("local %s={\n",config_name))

    local message_template = GetMessageTemplate(excel_template,message_name)
    if message_template ~= nil then
        for i=1,#message_template.var_list do
            -- #config 必然只有一个#var以及#type为一个自定义MessageType
            local message_var_string = MessageVarToLua(message_template.var_list[i],excel_template)
            if message_var_string ~= nil and #message_var_string > 0 then
                table.insert(string_builder,message_var_string)
            end
        end
    end
    table.insert(string_builder,string.format("}\n return %s\n",config_name))
    return table.concat(string_builder)
end

function MessageVarToLua(message_var,excel_template,row_data_target)
    local type = message_var.type
    local type_is_base = false
    for i=1, #var_base_type_list do
        if type == var_base_type_list[i] then
            type_is_base = true
            break
        else
            local var_base_list_type = "repeated "..var_base_type_list[i]
            if string.find(type,var_base_list_type) == 1 then
                type_is_base = true
                break
            end
        end
    end

    local message_var_string = nil
    if type_is_base then
        message_var_string = MessageBaseVarToLua(message_var,excel_template,row_data_target)
    else
        message_var_string = MessageTypeVarToLua(message_var,excel_template)
    end
    return message_var_string
end

function ToLuaTable(excel_template)
    local data_table = {}
    for key, value in pairs(excel_template.excel_config) do
        local lua_table = ConfigToLuaTable(key,value,excel_template)
        data_table[key] = {name=value,data = lua_table}
        print(lua_table)
    end
    return data_table
end