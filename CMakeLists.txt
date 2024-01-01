cmake_minimum_required(VERSION 3.22)

project (Staff-Allocation)

add_executable(Staff-Allocation Staff-Allocation.cpp)

include_directories( ${CMAKE_CURRENT_SOURCE_DIR}/libxl/include_cpp)

target_link_libraries(Staff-Allocation ${CMAKE_CURRENT_SOURCE_DIR}/libxl/lib64/libxl.so)