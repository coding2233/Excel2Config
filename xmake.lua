add_requires("lua") -- {system = false}
add_repositories("my-repo my-repositories")

-- add_requires("xlnt")
add_requires("xlnt",{configs = {shared = false}})



target("pb")
    set_kind("shared")
    add_files("lua-protobuf-master/pb.c")
    add_packages("lua")
    after_build(function (target) 
        local build_full_dir = "$(buildir)/$(plat)/$(arch)/$(mode)/"
        local target_name = target:name()
        local target_file_name = target:filename()
        local target_extension = path.extension(target_file_name)
        local package_path = build_full_dir.."package/"
        if not os.isdir(package_path) then 
            os.mkdir(package_path)
        end 
        --移动pb到对应的package
        os.mv(build_full_dir..target_file_name,package_path..target_name..target_extension)
        --复制lua
        os.cp("$(projectdir)/lua-protobuf-master/*.lua", package_path)
    end)

target("excel2config")
    -- set_kind("shared")
    set_languages("cxx17")
    add_files("src/*.cpp")
    -- add_includedirs("src")
    -- add_rules("utils.symbols.export_all", {export_classes = true})
    add_packages("xlnt","lua")
    -- add_deps("pb")
    after_build(function (target) 
        local build_full_dir = "$(buildir)/$(plat)/$(arch)/$(mode)/"
        local scripts_path = build_full_dir.."scripts/"
        if not os.isdir(scripts_path) then 
            os.mkdir(scripts_path)
        end 
        --复制lua
        os.cp("$(projectdir)/scripts/*.lua", scripts_path)
        local template_path = build_full_dir.."template/"
        if not os.isdir(template_path) then 
            os.mkdir(template_path)
        end 
        os.cp("$(projectdir)/template/*.xlsx", template_path)
    end)
