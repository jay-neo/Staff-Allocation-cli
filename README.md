 <div align='center'><h2>Staff Allocation</h2></div>


## How to Test the Program

### For Linux

- Step 1: Clone the repo in your local machine
```sh
git clone https://github.com/jay-neo/Staff-Allocation-cli.git
```

- Step 2: Run two commands in terminal (bash/zsh/fish)
```sh
cd Staff-Allocation-cli
```
```sh
sh build.sh
```

### For Windows

- Step 1: Clone the repo in your local machine
```sh
git clone https://github.com/jay-neo/Staff-Allocation-cli.git
```

- Step 2: Run two commands in terminal (pwsh)
```sh
Set-Location Staff-Allocation-cli
```
```sh
& (Join-Path (Get-Location) "build.ps1")
```


## Repository Structure
```
Staff-Allocation-cli
    │
    ├── libxl/
    │    ├── bin/
    │    │    └── libxl.dll
    │    ├── include_cpp/
    │    │       └── libxl.h (with other libxl header files)
    │    └─── lib64/
    │           ├── libxl.so
    │           └── libxl.dill
    │
    ├── CMakeLists.txt
    ├── main.cpp
    │
    └── Staff-Allocation.xlsx // This is the input file

```
