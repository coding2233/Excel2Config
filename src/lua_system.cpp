#include "lua_system.h"

namespace fs = std::filesystem;

static int GetFiles(lua_State* L)
{
    std::string path = lua_tostring(L,-1);
    lua_pop(L,1);

    lua_newtable(L);
    int file_index = 0;

    bool exists = fs::exists(path);
    if (exists)
    {
        if(fs::is_directory(path))
        {
            for (const auto &entry: fs::directory_iterator(path))
            {
                lua_pushnumber(L,++file_index);
                lua_pushstring(L,entry.path().c_str());
                lua_settable(L,-3);
            }

        }
        else
        {
            lua_pushnumber(L,++file_index);
            lua_pushstring(L,path.c_str());
            lua_settable(L,-3);
        }
    }
    
    return 1;
}


int luaopen_systemlib(lua_State *L)
{
    lua_register(L,"get_files",GetFiles);
    // luaL_newlib(L,systemlib);
    return 1;
}