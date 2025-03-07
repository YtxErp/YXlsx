# CMakeLists.txt for Console Application

cmake_minimum_required(VERSION 3.14)

project(HelloWorld LANGUAGES CXX)

set(CMAKE_CXX_STANDARD 20)

find_package(QT NAMES Qt6 COMPONENTS Core Gui REQUIRED)
find_package(Qt${QT_VERSION_MAJOR} COMPONENTS Core Gui REQUIRED)

add_subdirectory(YXlsx)

#########################
# Console Application {{

add_executable(HelloWorld main.cpp)
target_link_libraries(HelloWorld PRIVATE YXlsx::YXlsx)

# Console Application }}
########################
