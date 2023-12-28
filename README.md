
## Project Repository Structure
```
Staff-Allocation
    │
    ├── libxl/
    │    ├── bin/
    │    │    └── libxl.dll
    │    ├── include_cpp/
    │    │       └── libxl.h (with other libxl header files)
    │    ├──── lib/
    │    │      └── libxl.lib
    │    └─── CMakeLists.txt
    │
    ├── CMakeLists.txt
    └── main.cpp

```

## How to complie and run

### For Linux

- Step 1: Check **cmake** is installed or not
```sh
cmake --version
```

- Step 2: If not installed then
```sh
sudo apt install cmake
```
this will work for only Debian based Linux

- Step 3: Clone the repo in your local machine
```sh
git clone https://github.com/jay-neo/Staff-Allocation-cli.git
```

- Step 4: Copy and paste this code in terminal
```sh
cd Staff-Allocation-cli
cmake -B build
cd build
make

```


