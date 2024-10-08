extern "C"
{
    #include "lauxlib.h"
    #include "lualib.h"
    #include "lua.h"    
}

int main(int argc,char* args[])
{
    lua_State* L = luaL_newstate();
    luaL_openlibs(L);
    lua_close(L);
    
    return 0;
}