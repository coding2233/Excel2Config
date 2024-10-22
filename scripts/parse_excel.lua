-- require ("parse_excel_template")

local parse_excel = {}

local var_base_type_list = {"int32","string","bool","float"}

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
    local message_template = {type = type,row=kRow,cloumn = kCell, var_list = var_list}
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
        local var_string = message_var.var
        -- map<string,string> 
        -- repeated int32
        local sub_index = string.find(var_string,"#")
        if sub_index > 1 then
            var_string = string.sub(var_string,1,sub_index)
        end
        table.insert(string_builder, string.format("\t%s=%s;\n",var_string,tostring(i)))
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
    for i=1,#excel_template.enmu_list do
        local enmu_string = EnumToProtobuf(excel_template.enmu_list[i])
        if enmu_string~=nil and string.len(enmu_string) > 0 then
            table.insert(string_builder,enmu_string)
        end
    end

    -- MessageToProtobuf
    for i=1,#excel_template.message_template_list do
        local message_string = MessageToProtobuf(excel_template.message_template_list[i])
        if message_string~=nil and string.len(message_string) > 0 then
            table.insert(string_builder,message_string)
        end
    end

    local protobuf_string = table.concat(string_builder)
    print(protobuf_string)

end

local function GetMessageTemplate(excel_template,message_name)
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
    if message_template ~= nil then
        if is_list then
        else
            for i=1,#message_template.var_list do
                local message_var_string = MessageVarTemplteToLua(message_template.var_list[i],message_template,excel_template)
                if message_var_string ~= nil and #message_var_string > 0 then
                    table.insert(string_builder,message_var_string)
                end
            end
        end
    end

    return table.concat(string_builder)
end

function MessageMapVarToLua(message_var,excel_template)
    local string_builder = {}
    local type = message_var.type
    local var  = message_var.var
    local row = message_var.row
    local cloumn = message_var.cloumn
    
    if #string_builder > 0 then
        return table.concat(string_builder)
    else
        -- 继续处理message type类型
        return MessageTypeVarToLua(message_var,excel_template)
    end
end

function MessageBaseVarToLua(message_var,excel_template)
    local string_builder = {}
    local type = message_var.type
    local var  = message_var.var
    local row = message_var.row
    local cloumn = message_var.cloumn
    local type_string = nil
    local is_list = false
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
        local row_data = excel_template[row+2]
        if row_data ~= nil then
            table.insert(string_builder,string.format("%s=",var))
            local var_value = row_data[cloumn]
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
                    table.insert(string_builder,string.format("%s,",sub))
                    read_index = find_index + 1
                end
                table.insert(string_builder,"},\n")
            else
                table.insert(string_builder,string.format("%s,\n",var_value))
            end
        end
    end

    if #string_builder > 0 then
        return table.concat(string_builder)
    else
        -- 继续处理map类型
        return MessageMapVarToLua(message_var,excel_template)
    end
end

local function MessageVarTemplteToLua(message_var,excel_template)
    
    local string_builder = {}
     -- local var_data = {type = var_type, var = var_name, desc = var_desc}
    local type = message_var.type
    local var  = message_var.var 
    local row = message_var.row
    local cloumn = message_var.cloumn

    if type==nil or #type == 0 then
        return nil
    end


    -- -- map只支持基础类型
    -- if string.find(type,"map<") == 1 then
    -- -- list
    -- elseif string.find(type,"repeated ") == 1 then
    -- end

    -- 
    local var_string = MessageBaseVarToLua(message_var,excel_template)
    if var_string ~= nil and #var_string > 0 then
        table.insert(string_builder,var_string)
    end

    return table.concat(string_builder)
end

local function ConfigToLuaTable(message_name,config_name,excel_template)
    local string_builder = {}
    table.insert(string_builder,string.format("local %s={\n",config_name))

    local message_template = GetMessageTemplate(excel_template,message_name)
    if message_template ~= nil then
        for i=1,#message_template.var_list do
            local message_var_string = MessageVarTemplteToLua(message_template.var_list[i],message_template,excel_template)
            if message_var_string ~= nil and #message_var_string > 0 then
                table.insert(string_builder,message_var_string)
            end
        end
    end
    table.insert(string_builder,string.format("}\n return %s\n",config_name))
    return table.concat(string_builder)
end

function ToLuaTable(excel_template)
    for key, value in pairs(excel_template.excel_config) do
        local lua_table = ConfigToLuaTable(key,value,excel_template)
        print(lua_table)
    end
end