add_repositories("my-repo my-repositories")
add_requires("xlnt")

target("test")
    add_files("main.cpp")
    add_packages("xlnt")