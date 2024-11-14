VERSION = "0.0.1"

local exe_dir = get_exe_dir();
-- print(exe_dir)

package.path = exe_dir.."/package/?.lua;"..package.path
package.cpath = exe_dir.."/package/?.so;"..exe_dir.."/package/?.dylib;"..exe_dir.."/package/?.dll;"..package.cpath

-- print(package.path)
-- print(package.cpath)

-- function excel_cmd(cmd)
--     local index = cmd.find("=")
--     local len = string.len(cmd)
--     local arg = string.sub(cmd,index,len-index)
--     print(arg)
--     return {"excel"=arg}
-- end


local function help_cmd(cmd)
    return cmd
end

local function excel_cmd(cmd)
    local cmd_index = string.find(cmd,"=")
    if cmd_index == nil then
        return nil
    end

    if cmd_index >= string.len(cmd) then
        return nil
    end

    local cmd_arg = string.sub(cmd,cmd_index+1)
    return cmd_arg
end

local function out_cmd(cmd)
    local cmd_index = string.find(cmd,"=")
    if cmd_index == nil then
        return nil
    end

    if cmd_index >= string.len(cmd) then
        return nil
    end

    local cmd_arg = string.sub(cmd,cmd_index+1)
    return cmd_arg
end

local arg_configs = {
    {key="help", cmd={"--help","-h"},desc="print this text",callback=help_cmd},
    {key="version", cmd={"--version","-v"},desc=VERSION,callback=help_cmd},
    {key="excel", cmd={"--excel","-e"},desc="excel dir",callback=excel_cmd},
    {key="out", cmd={"--out","-o"},desc="out dir",callback=out_cmd},
}

local function parse_help_text()
    local help_text = {}
    for i = 1, #arg_configs do
        local arg_config = arg_configs[i]
        local cmds = arg_config.cmd
        local arg_text = "\t"
        for ci = 1, #cmds do
            arg_text=arg_text..cmds[ci]..","
        end
        arg_text = arg_text.."\t\t"..arg_config.desc
        table.insert(help_text,arg_text)
    end
    return help_text
end

function PrintHelp()
    local help_text = parse_help_text()
    for i = 1, #help_text do
        print(help_text[i])
    end
end

local function check_arg(arg)
    for a=1,#arg_configs do
        local arg_config = arg_configs[a]
        local key = arg_config.key
        local cmds = arg_config.cmd
        local callback = arg_config.callback
        for i=1,#cmds do
            local cmd = cmds[i]
            local find_cmd_index = string.find(arg,cmd)
            if find_cmd_index ~= nil and find_cmd_index == 1 then
                if callback ~= nil then
                    local result = callback(arg)
                    return key,result
                end
            end
        end
    end
end


function ParseArgs()
    local configs = {}
    for i=1,#args do
        local key,result = check_arg(args[i])
        if key ~= nil and result ~= nil then
            configs[key] = result
        end
    end
    return configs
end