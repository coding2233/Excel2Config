local exe_dir = get_exe_dir();
-- print(exe_dir)

package.path = exe_dir.."/package/?.lua;"..package.path
package.cpath = exe_dir.."/package/?.so;"..exe_dir.."/package/?.dylib;"..exe_dir.."/package/?.dll;"..package.cpath

-- print(package.path)
-- print(package.cpath)