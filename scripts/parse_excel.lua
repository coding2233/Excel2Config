local parse_excel = {}
local this = parse_excel

--基础类型
local var_base_type_list = {"int32","string","bool","float"}


function this.Load(excel_path)
    this.package = nil
    this.message_template_list = {}
    this.enmu_list = {}
    this.excel_config = {}
    --read excel
    local excel = read_excel(excel_path)
    for kSheet,vSheet in pairs(excel) do
        this.ParseExcel(vSheet)
    end
end



---ParseExcel [start]---

function this.ParseExcel(vSheet)
    for row = 1, #vSheet do
        local row_data = vSheet[row]
        for cloumn = 1, #row_data do
            local next_row = row
            local next_cloumn = cloumn
            local cell = row_data[cloumn]
            if "#message" == cell then
                next_row,next_cloumn = this.ParseMessage(vSheet,row,cloumn)
            -- 拼写错误
            elseif "#enum" == cell then
                next_row,next_cloumn = this.ParseEnum(vSheet,row,cloumn)
            elseif "#package" == cell then
                next_row,next_cloumn = this.ParsePackage(vSheet,row,cloumn)
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

function this.ParseMessage(vSheet,kRow,kCell)

    --检查
    if vSheet[kRow] == nil or vSheet[kRow+1] == nil or vSheet[kRow+2] == nil or vSheet[kRow+3] == nil then
        log.debug("excel message template cannot be parsed normally")
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
            log.debug("while var_list break")
            break
        end
        
        local var_name = vSheet[kRow+2][kCell+var_index]
        local var_desc = vSheet[kRow+3][kCell+var_index]
        local var_data = {type = var_type, var = var_name, desc = var_desc,row = kRow+2,cloumn = kCell+var_index}
        table.insert(var_list,var_data)
        var_index = var_index+1
    end

   
    -- 消息模板
    -- log.debug(type,kRow,kCell,#var_list)
    local message_template = {sheet=vSheet,type = type,row=kRow,cloumn = kCell, var_list = var_list}
    table.insert(this.message_template_list,message_template)

    -- 配置入口
    if vSheet[kRow-1] ~= nil and vSheet[kRow-1][kCell] ~= nil and "#config" == vSheet[kRow-1][kCell] then
        local config_value = vSheet[kRow-1][kCell+1]
        if config_value ~= nil and  #config_value > 0 then
            this.excel_config[type] = config_value
            log.debug("excel_config",type,config_value,#config_value,kRow,kCell)
        end
    end

    return kRow,kCell
end

function this.ParseEnum(vSheet,kRow,kCell)
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
    table.insert(this.enmu_list,enum_type)
    return kRow+enum_value_index,1
end

function this.ParsePackage(vSheet,kRow,kCell)
    this.package = vSheet[kRow][kCell+1]
    log.debug("ParsePackage",this.package)

    return kRow,kCell+1
end

---ParseExcel [end]---




---ToProtobuf [start]---

function this.ToProtobuf()
    log.debug("ToProtobuf")
    local string_builder = {}
    local string_package = nil
    -- table.insert(string_builder,"syntax = \"proto3\";\n\n")
    if this.package ~= nil and string.len(this.package) > 0 then
        string_package = string.format("package %s;\n\n",this.package)
    end

    -- EnumToProtobuf
    for i=1,#this.enmu_list do
        local enmu_string = this.EnumToProtobuf(this.enmu_list[i])
        if enmu_string~=nil and string.len(enmu_string) > 0 then
            table.insert(string_builder,enmu_string)
        end
    end

    -- MessageToProtobuf
    for i=1,#this.message_template_list do
        local message_string = this.MessageToProtobuf(this.message_template_list[i])
        if message_string~=nil and #message_string > 0 then
            -- log.debug(message_string)
            table.insert(string_builder,message_string)
        end
    end

    --log.debug(#string_builder)
    local protobuf_string = table.concat(string_builder)
    log.debug("ToProtobuf\n",protobuf_string)
    return protobuf_string, string_package
end

function this.MessageToProtobuf(message_template)
    log.debug("message_template",message_template,message_template.type,#message_template.var_list)
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
        -- log.debug("--------",var_type_string,sub_index)
        if sub_index ~= nil and sub_index > 1 then
            var_type_string = string.sub(var_type_string,1,sub_index-1)
        end
        table.insert(string_builder, string.format("\t%s %s = %s;\n",var_type_string,var_string,tostring(i)))
    end
    table.insert(string_builder, "}\n\n")

    return table.concat(string_builder)
end

function this.EnumToProtobuf(enum_template)
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
        table.insert(string_builder, string.format("\t%s:%s;\n",enum_value.var,enum_value.value))
    end
    table.insert(string_builder, "}\n\n")

    return table.concat(string_builder)
end

---ToProtobuf [end]---


---ToJson [start]---

function this.ToJson()
    local data_table = {}
    for key, value in pairs(this.excel_config) do
        log.debug("ToLuaTable",key,value)
        local config_json,config_textproto = this.ConfigToJson(key,value)
        config_json = string.gsub(config_json,",}","}")
        config_json = string.gsub(config_json,",]","]")
        config_textproto = string.gsub(config_textproto,",}","}")
        config_textproto = string.gsub(config_textproto,",]","]")
        data_table[key] = {name=value,json = config_json,textproto = config_textproto}
        log.debug(string.format("ToLuaTable\n%s\n%s\n%s\n\n%s\n",key,value,config_json,config_textproto))
    end
    return data_table
end

function this.ConfigToJson(message_name,config_name)
    log.debug("ConfigToJson",message_name,config_name)
    local json_string_builder = {}
    local textproto_string_builder = {}
    table.insert(json_string_builder,string.format("{",config_name))

    local message_template = this.GetMessageTemplate(message_name)
    if message_template ~= nil then
        local row_data = message_template.sheet[message_template.row+4]
        for i=1,#message_template.var_list do
            local json_message_var_string,textproto_message_var_string = this.MessageVarToLua(message_template.var_list[i],row_data)
            if json_message_var_string ~= nil and #json_message_var_string > 0 then
                table.insert(json_string_builder,json_message_var_string)
                table.insert(json_string_builder,",")
            end
            if textproto_message_var_string ~= nil and #textproto_message_var_string > 0 then
                table.insert(textproto_string_builder,textproto_message_var_string)
                table.insert(textproto_string_builder,",")
            end
        end
    end
    table.insert(json_string_builder,string.format("}",config_name))
    return table.concat(json_string_builder),table.concat(textproto_string_builder)
end

function this.GetMessageTemplate(message_name)
    log.debug("GetMessageTemplate",message_name,#this.message_template_list)
    for i=1,#this.message_template_list do
        local message_template = this.message_template_list[i]
        if message_name == message_template.type then
            log.debug("GetMessageTemplate found",message_name,#this.message_template_list,message_template.type,message_template)
            return message_template
        end
    end
    log.error("GetMessageTemplate not found.",message_name)
    return nil
end

function this.MessageVarToLua(message_var,row_data_target)
    log.debug("MessageVarToLua",message_var, message_var.type,row_data_target)
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

    log.debug("MessageVarToLua",type,type_is_base)
    local json_message_var_string = nil
    local textproto_message_var_string = nil
    if type_is_base then
        json_message_var_string,textproto_message_var_string = this.MessageBaseVarToLua(message_var,row_data_target)
    else
        json_message_var_string,textproto_message_var_string = this.MessageTypeVarToLua(message_var)
    end
    return json_message_var_string,textproto_message_var_string
end

function this.MessageTypeVarToLua(message_var)
    local json_string_builder = {}
    local textproto_string_builder = {}
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

    local message_template = this.GetMessageTemplate(type_name)
    log.debug("MessageTypeVarToLua",type,type_name,is_list,var,message_template)

    if message_template ~= nil then
        -- for key, value in pairs(message_template) do
        --     log.debug(key,value)
        -- end
        local find_message_template_row = message_template.row
        log.debug( message_template.row,message_template.cloumn, message_template.type,message_template.sheet)
        local val_symbol_start = "{"
        local val_symbol_end = "}"
        if is_list then
            val_symbol_start = "["
            val_symbol_end = "]"
        end
        table.insert(json_string_builder,string.format("\"%s\":%s",var,val_symbol_start))
        table.insert(textproto_string_builder,string.format("%s:%s",var,val_symbol_start))
        if is_list then
            local find_row_index = 4;
            while true do
                local find_row_data = message_template.sheet[find_message_template_row+find_row_index]
                log.debug(find_message_template_row,find_row_index,find_row_data)
                if find_row_data ~= nil then
                    -- todo: 检查这一行都是空数据 100太假了
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
                    table.insert(json_string_builder,"{")
                    table.insert(textproto_string_builder,"{")
                    -- log.debug(#message_template.var_list)
                    for i=1,#message_template.var_list do
                        local message_var_string = this.MessageVarToLua(message_template.var_list[i],find_row_data)
                        if message_var_string ~= nil and #message_var_string > 0 then
                            table.insert(json_string_builder,string.format("%s,",message_var_string))
                            table.insert(textproto_string_builder,string.format("%s,",message_var_string))
                        end
                    end
                    table.insert(json_string_builder,"},")
                    table.insert(textproto_string_builder,"},")
                else
                    -- 空行数据
                    break
                end
                find_row_index=find_row_index+1
            end
            
        else
            local find_row_data = message_template.sheet[find_message_template_row+4]

            for i=1,#message_template.var_list do
                local message_var_string = this.MessageVarToLua(message_template.var_list[i],find_row_data)
                if message_var_string ~= nil and #message_var_string > 0 then
                    table.insert(json_string_builder,message_var_string)
                    table.insert(textproto_string_builder,message_var_string)
                end
            end
        end
        table.insert(json_string_builder,string.format("%s,",val_symbol_end))
        table.insert(textproto_string_builder,string.format("%s,",val_symbol_end))
    end

    return this.TableConcatEx(json_string_builder),this.TableConcatEx(textproto_string_builder)
end

function this.MessageBaseVarToLua(message_var,row_data)
    local json_string_builder = {}
    local textproto_string_builder = {}
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

    -- log.debug("MessageBaseVarToLua",type_string,var,row,cloumn,row_data)
    if type_string ~= nil then
        local find_type_is_string = string.find(type_string,"string")
        type_is_string = find_type_is_string ~= nil and find_type_is_string >0

        local find_type_is_bool = string.find(type_string,"bool")
        type_is_bool = find_type_is_bool ~= nil and find_type_is_bool > 0

        -- local row_data = row_data_target
        -- if row_data == nil then
        --     row_data = excel_template[row+2]
        -- end
        
        if row_data ~= nil then
            local var_value = row_data[cloumn]
            -- 需要处理一下默认数据
            if var_value == nil or #var_value==0 then
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
            log.debug("MessageBaseVarToLua",type_string,var,row,cloumn,row_data,var_value)
            table.insert(json_string_builder,string.format("\"%s\":",var))
            table.insert(textproto_string_builder,string.format("%s:",var))
            if is_list then
                -- todo ..  
                -- e.g.
                -- repeated int32#sep=,
                local find_key = ","
                table.insert(json_string_builder,"[")
                table.insert(textproto_string_builder,"[")
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

                    table.insert(json_string_builder,string.format(var_vale_format,sub))
                    table.insert(textproto_string_builder,string.format(var_vale_format,sub))
                    read_index = find_index + 1
                end
                table.insert(json_string_builder,"],")
                table.insert(textproto_string_builder,"],")
            else
                local var_vale_format = "%s,"
                if type_is_string then
                    var_vale_format = "\"%s\","
                end
                if type_is_bool then
                    var_value = string.lower(var_value)
                end
                table.insert(json_string_builder,string.format(var_vale_format,var_value))
                table.insert(textproto_string_builder,string.format(var_vale_format,var_value))
            end
        end
    end

    if #json_string_builder > 0 then
        return this.TableConcatEx(json_string_builder),this.TableConcatEx(textproto_string_builder)
    else
        -- 继续处理map类型
        local json_value,textproto_value this.MessageMapVarToLua(message_var,row_data)
        return json_value,textproto_value
    end
end

function this.MessageMapVarToLua(message_var,row_data_target)
    local json_string_builder = {}
    local textproto_string_builder = {}
    local type = message_var.type
    local var  = message_var.var
    local row = message_var.row
    local cloumn = message_var.cloumn
    
    local var_value = row_data_target[cloumn]
    if var_value == nil then
        var_value = ""
    end

    table.insert(json_string_builder,"{")
    -- todo...检查value string类型
    for k, v in string.gmatch(var_value, "(%w+):(%w+)") do
        table.insert(json_string_builder,string.format("\"%s\":%s,",k,v))
        if #textproto_string_builder > 0 then
            table.insert(textproto_string_builder,string.format("%s:,",var))
        end
        table.insert(textproto_string_builder,string.format("{key:%s\nvalue:%s},",k,v))
    end
    table.insert(json_string_builder,"},")

    return this.TableConcatEx(json_string_builder),this.TableConcatEx(textproto_string_builder)
end

function this.TableConcatEx(table_data)
    local string_builder_result = table.concat(table_data)
    local i,result_j = string.find(string_builder_result,",",-1)
    if result_j == #string_builder_result then
        string_builder_result = string.sub(string_builder_result,1,result_j-1)
    end
    return string_builder_result
end
---ToLuaTable [end]---

return parse_excel