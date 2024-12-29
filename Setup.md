# Setup

## Import in CMakeLists

Two possible ways

- first

```CMakeLists
set(YXLSX_PATH ${CMAKE_CURRENT_SOURCE_DIR}/YXlsx/YXlsx)

if(NOT EXISTS ${YXLSX_PATH})
    message(FATAL_ERROR
        "YXlsx library not found at ${YXLSX_PATH}. Please download it from:\n"
        "https://github.com/YtxErp/YXlsx.git\n"
        "and ensure the library is placed under ${CMAKE_CURRENT_SOURCE_DIR}/"
    )
endif()

add_subdirectory(${YXLSX_PATH})

target_link_libraries(app PRIVATE YXlsx::YXlsx)
```

- second

```CMakeLists
include(FetchContent)

FetchContent_Declare(
  YXlsx
  GIT_REPOSITORY https://github.com/YtxErp/YXlsx.git
  GIT_TAG        sha-of-the-commit
  SOURCE_SUBDIR  YXlsx
)
FetchContent_MakeAvailable(YXlsx)
target_link_libraries(myapp PRIVATE YXlsx::YXlsx)
```
