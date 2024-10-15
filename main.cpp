

extern "C"
{
    #include "lauxlib.h"
    #include "lualib.h"
    #include "lua.h"    
}

#include <iostream>
#include <xlnt/xlnt.hpp>

static const char* global_exe_dir;

static int get_exe_dir(lua_State* L)
{
    // printf("get_exe_dir:%s\n",global_exe_dir);
    lua_pushstring(L, global_exe_dir);
    return 1;
} 

static int show_stack_count(lua_State* L,const char* test_key)
{
    int count = lua_gettop(L);
    printf("[%s] show_stack_count: %d\n",test_key,count);
    return count;
}

//参考: https://cloud.tencent.com/developer/ask/sof/110316520/answer/122621096
static int read_excel(lua_State* L)
{
    std::string excel_path = lua_tostring(L,-1);
    lua_pop(L,1);

    xlnt::workbook wb;
    wb.load(excel_path);

    std::size_t sheet_count = wb.sheet_count();

    if ( sheet_count > 0)
    {
        //excel table
         lua_newtable(L);
        //lua_createtable(L, sheet_count,0);

        for (size_t i = 0; i < sheet_count; i++)
        {
            auto ws = wb.sheet_by_index(i);
            auto ws_title = ws.title();
            // printf("%s\n",ws_title.c_str());

            //设置excel table的sheet key
            //lua_pushstring(L,ws_title.c_str());
            //创建sheet table
            lua_newtable(L);
            
            //读取sheet里面的数据
            //空的cell不被跳过，保证数据的完整
            int row_index = 0;
            for (auto row : ws.rows(false))
            {
                //行table
                lua_pushnumber(L,++row_index);

                // lua_pushstring(L,"stack test data中文");

                lua_newtable(L);

                int colum_index = 0;
                for (auto cell : row)
                {
                    lua_pushnumber(L,++colum_index);
                    lua_pushstring(L,cell.to_string().c_str());
                    //设置列数据 -> row table  id作为key
                    lua_settable(L, -3);
                    // std::cout<< cell.to_string().c_str()<<std::endl;
                }

                //设置当前行的table -> sheet table, id作为key
                lua_settable(L, -3);
            }

            /* Remember, child table is on-top of the stack.
            * lua_settable() pops key, value pair from Lua VM stack. */
            //lua_settable(L, -3);
            //设置sheet table -> excel table，并以sheet name作为key
            lua_setfield(L,-2,ws_title.c_str());
        }

        return 1;
    }
    
    return 0;
}

static int load_excel(lua_State* L)
{
    xlnt::workbook wb;
    xlnt::worksheet ws = wb.active_sheet();
    lua_pushlightuserdata(L,&ws);
    lua_pushlightuserdata(L,&wb);
    return 2;
}

static int write_worksheet(lua_State* L)
{
    xlnt::worksheet* ws = (xlnt::worksheet* )lua_touserdata(L,-1);
    int row = lua_tonumber(L,-2);
    int column = lua_tonumber(L,-3);
    std::string value = lua_tostring(L,-4);

    auto cell = ws->cell(row,column);
    cell.value(value);
    lua_pop(L,4);
    return 0;
}

static int save_excel(lua_State* L)
{
    xlnt::workbook* wb = (xlnt::workbook* )lua_touserdata(L,-1);
    std::string excel_name = lua_tostring(L,-2);
    wb->save(excel_name.c_str());
    lua_pop(L,2);
    return 0;
}


int setLuaPath( lua_State* L, const char* path )
{
    lua_getglobal( L, "package" );
    lua_getfield( L, -1, "path" ); // get field "path" from table at top of stack (-1)
    std::string cur_path = lua_tostring( L, -1 ); // grab path string from top of stack
    cur_path.append( ";" ); // do your path magic here
    cur_path.append( path );
    lua_pop( L, 1 ); // get rid of the string on the stack we just pushed on line 5
    lua_pushstring( L, cur_path.c_str() ); // push the new one
    lua_setfield( L, -2, "path" ); // set the field "path" in table at -2 with value at top of stack
    lua_pop( L, 1 ); // get rid of package table from top of stack
    return 0; // all done!
}

int main(int argc,char* args[])
{
    std::string exe_path = std::string(args[0]);
    int subCount = exe_path.find_last_of('\\');
    if (subCount <= 0)
    {
        subCount =  exe_path.find_last_of('/');
    }
    std::string exe_dir = exe_path.substr(0,subCount);
    global_exe_dir = exe_dir.c_str();
    std::string lua_scripts_dir = std::string(exe_dir).append("/scripts");
    std::string lua_path = std::string(lua_scripts_dir).append("/main.lua");
    lua_scripts_dir.append("/?.lua;");
    // printf("path:%s\n dir:%s\n lua:%s \n\n",exe_path.c_str(),exe_dir.c_str(),lua_path.c_str());

    lua_State* L = luaL_newstate();
    luaL_openlibs(L);

    lua_register(L,"get_exe_dir",get_exe_dir);

    lua_register(L,"read_excel",read_excel);
    lua_register(L,"load_excel",load_excel);
    lua_register(L,"save_excel",save_excel);
    lua_register(L,"write_worksheet",save_excel);


    //设置scripts/main.lua的package环境，以及运行入口
    setLuaPath(L,lua_scripts_dir.c_str());
    luaL_dofile(L,lua_path.c_str());
    

    lua_close(L);
    
    // xlnt::workbook wb;
    // xlnt::worksheet ws = wb.active_sheet();

    // ws.cell("A1").value(5);
    // ws.cell("B2").value("string data");
    // wb.save("example.xlsx");

    // xlnt::workbook wb;
    // wb.load("example.xlsx");

    // auto ws = wb.active_sheet();
    // for (auto row : ws.rows(true))
    // {
    //     for (auto cell : row)
    //     {
    //         std::cout<< cell.to_string()<<std::endl;
    //     }
        
    // }
    

    return 0;
}

//lua加载c动态库
//int luaopen_函数名库名称(lua_State* L)
//
//https://blog.csdn.net/Xiao_brother/article/details/127210803