#ifndef __LUA_SYSTEM_H__
#define __LUA_SYSTEM_H__

extern "C"
{
    #include "lauxlib.h"
    #include "lualib.h"
    #include "lua.h"    
}

#include <iostream>
#include <filesystem>
#include <vector>

int luaopen_systemlib(lua_State *L);


#endif