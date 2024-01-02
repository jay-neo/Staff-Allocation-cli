cmake_minimum_required(VERSION 3.22)

project (Staff-Allocation)

# set(CMAKE_CXX_STANDARD 17)

set(CMAKE_RUNTIME_OUTPUT_DIRECTORY ${CMAKE_CURRENT_SOURCE_DIR}/bin)

add_executable(Staff-Allocation Staff-Allocation.cpp)

include_directories( ${CMAKE_CURRENT_SOURCE_DIR}/libxl/include_cpp)

if(WIN32)
    target_link_libraries(Staff-Allocation ${CMAKE_CURRENT_SOURCE_DIR}/libxl/lib64/libxl.dll)
else()
    target_link_libraries(Staff-Allocation ${CMAKE_CURRENT_SOURCE_DIR}/libxl/lib64/libxl.so)
endif()