add_repositories("my-repo my-repositories")
add_requires("xlnt","lua")

target("test")
    -- set_kind("shared")
    add_files("main.cpp")
    -- add_rules("utils.symbols.export_all", {export_classes = true})
    add_packages("xlnt","lua")