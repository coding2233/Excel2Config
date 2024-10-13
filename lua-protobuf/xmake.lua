add_requires("lua",{configs = {shared = true}})

target("pb")
    set_kind("shared")
    add_files("lua-protobuf-master/pb.c")
    add_packages("lua")