cmake_minimum_required(VERSION 3.10)
project(FilterDataProject) #项目名

set(CMAKE_CXX_STANDARD 14) #C++标准

#关键：设置可执行文件输出到项目根目录的bin文件夹
set(CMAKE_RUNTIME_OUTPUT_DIRECTORY ${CMAKE_SOURCE_DIR}/bin)

# 头文件路径
include_directories(
    ./include
    ./xlnt-1.5.0/include
)

# 库文件路径
link_directories(
    ./xlnt-1.5.0/build/source
)

# 所有源文件（main.cc + src下的所有.cc）
file(GLOB SRCS 
    main.cc
    src/*.cc
)

# 生成可执行文件
add_executable(main ${SRCS})

# 链接第三方库
target_link_libraries(main xlntd)