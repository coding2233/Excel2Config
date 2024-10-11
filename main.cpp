extern "C"
{
    #include "lauxlib.h"
    #include "lualib.h"
    #include "lua.h"    
}

#include <iostream>
#include <xlnt/xlnt.hpp>

int main(int argc,char* args[])
{
    lua_State* L = luaL_newstate();
    luaL_openlibs(L);
    lua_close(L);
    
    // xlnt::workbook wb;
    // xlnt::worksheet ws = wb.active_sheet();

    // ws.cell("A1").value(5);
    // ws.cell("B2").value("string data");
    // wb.save("example.xlsx");

    xlnt::workbook wb;
    wb.load("example.xlsx");

    auto ws = wb.active_sheet();
    for (auto row : ws.rows(true))
    {
        for (auto cell : row)
        {
            std::cout<< cell.to_string()<<std::endl;
        }
        
    }
    

    return 0;
}