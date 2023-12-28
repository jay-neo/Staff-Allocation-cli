# CMakeLists.txt in the main directory

cmake_minimum_required(VERSION 3.10)

project(StaffAllocation)

# Set the C++ standard to use
set(CMAKE_CXX_STANDARD 17)
set(CMAKE_CXX_STANDARD_REQUIRED ON)

# Define the source files
set(SOURCE_FILES
    main.cpp
)

# Set the include directories
include_directories(
    ${CMAKE_CURRENT_SOURCE_DIR}/libxl/include_cpp
)

# Add subdirectories
add_subdirectory(libxl)

# Create the executable
add_executable(StaffAllocation ${SOURCE_FILES})

# Link the executable with libxl_target
target_link_libraries(StaffAllocation libxl_target stdc++fs)
