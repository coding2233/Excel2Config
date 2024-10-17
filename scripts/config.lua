local exe_dir = get_exe_dir();
-- print(exe_dir)

package.path = exe_dir.."/package/?.lua;"..package.path
package.cpath = exe_dir.."/package/?.so;"..exe_dir.."/package/?.dylib;"..exe_dir.."/package/?.dll;"..package.cpath

-- print(package.path)
-- print(package.cpath)

function help_cmd(cmd)
    print("--help print this text")
end

local arg_configs = {
    {cmd={"--help","-h"},desc="print this text",callback=help_cmd},
    -- {"cmd"={"--help","-h"},"desc"="print this text","callback"=help_cmd}
    -- {"cmd"={"--excel"},"desc"="excel path","callback"=excel_cmd}
}

function parse_args()
    local configs = {}
    for i=1,#args do
        local result = check_arg(args[i])
        if result ~= nil then
            table.insert(configs,result)
        end
    end
    return configs
end


function check_arg(arg)
    for a=1,#arg_configs do
        local arg_config = arg_configs[a]
        local cmds = arg_config.cmd
        local callback = arg_config.callback
        for i=1,#cmds do
            local cmd = cmds[i]
            local find_cmd_index = string.find(arg,cmd)
            if find_cmd_index ~= nil and find_cmd_index == 1 then
                print("find_cmd_index"..tostring(find_cmd_index))
                if callback ~= nil then
                    print("find_cmd_index"..tostring(find_cmd_index))
                    local result = callback(cmd)
                end
            end
        end
    end
end




-- function excel_cmd(cmd)
--     local index = cmd.find("=")
--     local len = string.len(cmd)
--     local arg = string.sub(cmd,index,len-index)
--     print(arg)
--     return {"excel"=arg}
-- end