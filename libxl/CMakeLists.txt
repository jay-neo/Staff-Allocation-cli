# CMakeLists.txt in the libxl directory

# Define the library location
set(LIBXL_LIBRARY_PATH ${CMAKE_CURRENT_SOURCE_DIR}/lib/libxl.lib)

# Create a custom target for libxl
add_library(libxl_target STATIC IMPORTED GLOBAL)
set_target_properties(libxl_target PROPERTIES
    IMPORTED_LOCATION ${LIBXL_LIBRARY_PATH}
)

# Add include directories for libxl
target_include_directories(libxl_target
    INTERFACE
    ${CMAKE_CURRENT_SOURCE_DIR}/include_cpp
)

# Optionally, you can install the library if needed
install(FILES ${LIBXL_LIBRARY_PATH} DESTINATION lib)
