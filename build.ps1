if (-not (Test-Path (Get-Command cmake.exe -ErrorAction SilentlyContinue).Path)) {
    Write-Host "    ##################################"
    Write-Host "    #     CMake is not installed.    #"
    Write-Host "    ##################################"
    
    $choice = Read-Host "    Do you want to install it? (y/n)"
    
    if ($choice -eq 'y' -or $choice -eq 'Y') {
        Write-Host "Installing CMake..."
        
        if (Get-Command choco -ErrorAction SilentlyContinue) {
            choco install cmake -y
        }
        elseif (Test-Path (Get-Command winget -ErrorAction SilentlyContinue)) {
            winget install cmake
        }
        else {
            Write-Host "    Unsupported package manager. Please install CMake manually."
            Write-Host "    Before running this program, make sure CMake is installed on your machine."
            exit 1
        }

        Write-Host "CMake has been installed."
    }
    else {
        Write-Host "    Before running this program, make sure CMake is installed on your machine."
    }
}

if (Test-Path -Path ".\Build" -PathType Container) {
    Write-Host "'Build' folder is already present!"
    
    $choice = Read-Host "Do you want to delete it? (y/n)"
    if ($choice -eq 'y' -or $choice -eq 'Y') {
        Write-Host "Deleting existing 'Build' folder..."
        Remove-Item -Path ".\Build" -Recurse -Force
    } else {
        return
    }
}


mkdir build
cd build
cmake -G "MinGW Makefiles" ..
mingw32-make
cd ..
cp libxl/lib64/libxl.dll bin/
cp Staff-Allocation.xlsx bin/
.\bin\Staff-Allocation.exe
