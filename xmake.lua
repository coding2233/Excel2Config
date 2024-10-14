add_requires("lua",{system = false})
add_repositories("my-repo my-repositories")

-- add_requires("xlnt")
add_requires("xlnt",{configs = {shared = false}})


target("pb")
    set_kind("shared")
    add_files("lua-protobuf-master/pb.c")

target("xlnt2config")
    -- set_kind("shared")
    add_files("main.cpp")
    -- add_rules("utils.symbols.export_all", {export_classes = true})
    set_languages("cxx17")
    add_packages("xlnt","lua")
    -- add_deps("pb")
