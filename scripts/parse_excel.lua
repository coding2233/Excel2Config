-- require ("excel")

local package = nil
local message_list = {}
local enmu_list = {}
local excel_config = {}
local message_template_list = {}

local function ParseMessage(vSheet,kRow,kCell)
    local type = vSheet[kRow][kCell+1]
    if type == nil or string.len(type) == 0 or #type==0 then
        return
    end

    -- 消息模板
    local message_template = {type = type,row=kRow,cloumn = kCell}
    table.insert(message_template_list,message_template)

    -- 配置入口
    if vSheet[kRow-1][kCell] ~= nil and #"config" == vSheet[kRow-1][kCell] then
        local config_value = vSheet[kRow-1][kCell+1]
        if config_value ~= nil then
            excel_config[type] = config_value
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
    table.insert(enmu_list,enum_type)
    return kRow+enum_value_index,0
end

local function ParsePackage(vSheet,kRow,kCell)
    package = vSheet[kRow][kCell+1]
    print("ParsePackage",package)

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
    print("ParseExcel",path)
    local files = get_files(path)
    for i = 1,#files do
        ParseExcelFile(files[i])
    end
end

