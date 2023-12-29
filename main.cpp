#include <iostream>
#include <cstring>
#include <vector>
#include <algorithm>
#include <random>
#include <experimental/filesystem>
#include <windows.h>
#include "libxl.h"


std::string ConvertTCHARToString(const TCHAR* tcharString) {
#ifdef UNICODE
    // Convert wide string to multi-byte string
    int bufferSize = WideCharToMultiByte(CP_UTF8, 0, tcharString, -1, nullptr, 0, nullptr, nullptr);
    std::string result(bufferSize, 0);
    WideCharToMultiByte(CP_UTF8, 0, tcharString, -1, &result[0], bufferSize, nullptr, nullptr);
#else
    // If TCHAR is char, no conversion is needed
    std::string result(tcharString);
#endif
    return result;
}


const TCHAR* ConvertStringToTCHAR(const std::string& str) {
#ifdef UNICODE
    // Convert multi-byte string to wide string
    int bufferSize = MultiByteToWideChar(CP_UTF8, 0, str.c_str(), -1, nullptr, 0);
    wchar_t* result = new wchar_t[bufferSize];
    MultiByteToWideChar(CP_UTF8, 0, str.c_str(), -1, result, bufferSize);
#else
    // If TCHAR is char, no conversion is needed
    const TCHAR* result = str.c_str();
#endif
    return result;
}



struct Staff {
    std::vector<std::pair<int, std::string>> Job;
    int total;
    int recent;
    Staff() {
        total = 0;
        recent = 0;
    }
    void addJob(int sft, std::string w) {
        recent = 1;
        total += 1;
        Job.emplace_back(sft, w);
    }
    void addJob() {
        recent = 0;
    }
};

class StaffAllocation {
    int st;
    libxl::Book* givenFile = nullptr;
    libxl::Book* staffFile = nullptr;
    std::string inputFileName;
    std::string outputFileName = "Staff Allocation Sheet.xlsx";
    std::vector<std::pair<std::string, struct Staff>> staffs;
    std::vector<std::string> rooms;

public:
    StaffAllocation(std::string inputFilePath, int _st = 2) : inputFileName(inputFilePath), st(_st) {
        givenFile = xlCreateXMLBookW();
        staffFile = xlCreateXMLBookW();
    }

    ~StaffAllocation() {
        givenFile->release();
        staffFile->release();
    }

    void shuffle(const std::string& s) {
        std::random_device rd;
        std::mt19937 g(rd());
        if (s == "staffs") {
            std::shuffle(staffs.begin(), staffs.end(), g);
        }
        else if (s == "rooms") {
            std::shuffle(rooms.begin(), rooms.end(), g);
        }
    }

    void sortStaff() {
        std::sort(staffs.begin(), staffs.end(), [](const auto& lhs, const auto& rhs) {
            if (lhs.second.total != rhs.second.total) { // error
                return lhs.second.total < rhs.second.total;
            } else if (lhs.second.recent != rhs.second.recent) {
                return lhs.second.recent < rhs.second.recent;
            } else {
                return lhs.first < rhs.first;
            }
        });
    }


    void allocateSheet(const int& sheetIdx) {
        libxl::Sheet* givenSheet = givenFile->getSheet(sheetIdx);
        libxl::Sheet* staffSheet = staffFile->addSheet(givenSheet->name());

        int Rows = givenSheet->lastRow(); // +1
        int Columns = 0; //+1
        int itr = -1;


        for (int r = 0; r < givenSheet->lastCol(); ++r) {
            if (givenSheet->cellType(0, r) != libxl::CELLTYPE_BLANK) {
                ++Columns;
            }
        }


        staffSheet->writeStr(0, 0, givenSheet->readStr(0, 0));

        for (int r = 1; r < Rows; ++r) {
            if (givenSheet->cellType(r, 0) != libxl::CELLTYPE_BLANK) {
                staffs.emplace_back(ConvertTCHARToString(givenSheet->readStr(r, 0)), Staff());
                staffSheet->writeStr(r, 0, givenSheet->readStr(r, 0));
                ++itr;
            }
        }
        int totalStaff = itr;

        this->shuffle("staffs");

        for (int shift = 1; shift < Columns; ++shift) {
            // after first shift sort the 
            if (shift != 1) {
                this->sortStaff();
            }

            // calculate room number for each shift
            int rows = 1;
            for (int r = 1; r < Rows; ++r) {
                if (givenSheet->cellType(r, shift) != libxl::CELLTYPE_BLANK) {
                    ++rows;
                    rooms.emplace_back(ConvertTCHARToString(givenSheet->readStr(r, shift)));
                }
            }

            this->shuffle("rooms");
            staffSheet->writeStr(0, shift, givenSheet->readStr(0, shift));

            // Final allocate for each shift
            itr = 0;
            for (int r = 0; r < rows; ++r) {
                for (int i = 0; i < st; ++i) {
                    staffs[itr++].second.addJob(shift, rooms[r]);
                }
            }

            // those not hired, those (recent = 0)
            while (itr <= totalStaff) {
                staffs[itr++].second.addJob();
            }

            rooms.clear();

        }
        for (itr = 0; itr < totalStaff; ++itr) {
            for (int i = 0; i < staffs[itr].second.Job.size(); ++i) {
                staffSheet->writeStr(itr + 1, staffs[itr].second.Job[i].first, ConvertStringToTCHAR(staffs[itr].second.Job[i].second));
            }
        }

        staffs.clear();
        

    }

    void neo() {
        if (givenFile->load(ConvertStringToTCHAR(inputFileName))) {
            
            this->allocateSheet(0);
            /*for (int sheetIdx = 0; sheetIdx < ivenFile->sheetCount(); ++sheetIdx) {
                this->allocateSheet(sheetIdx);
            }*/
            staffFile->save(ConvertStringToTCHAR(outputFileName));

        }
        else {
            std::cerr << "Input Excel File not found !!" << std::endl;
            std::cout << "Press any key to exit...";
            std::cin.get();
        }
    }
};



bool welcomeMsg() {

    std::string msg = "no";
    std::cout << "Prerequisites:" << std::endl;
    std::cout << "1. Excel file(.xlsx) should be formatted like this ??" << std::endl;
    std::cout << "???????????????????????????????????????????????????????" << std::endl;
    std::cout << "? Staff Name  ? Shift 1 ? Shift 2 ? Shift 3 ? Shift n ?" << std::endl;
    std::cout << "???????????????????????????????????????????????????????" << std::endl;
    std::cout << "? Jagadish S  ?   A100  ?   A101  ?   A100  ?   B100  ?" << std::endl;
    std::cout << "???????????????????????????????????????????????????????" << std::endl;
    std::cout << "?  Subarna M  ?   A101  ?   B100  ?         ?   B101  ?" << std::endl;
    std::cout << "???????????????????????????????????????????????????????" << std::endl;
    std::cout << "? Soumyajit N ?         ?   B101  ?         ?         ?" << std::endl;
    std::cout << "???????????????????????????????????????????????????????" << std::endl;
    std::cout << "?  Sourita C  ?         ?         ?         ?         ?" << std::endl;
    std::cout << "???????????????????????????????????????????????????????" << std::endl;
    std::cout << "?  person m   ?         ?         ?         ?         ?" << std::endl;
    std::cout << "???????????????????????????????????????????????????????" << std::endl;
    std::cout << "2. The excel file name should be same name as the execution file(.exe)" << std::endl;

    while (msg != "jay-neo") {
        std::cout << "If you understand the prerequisites and continue the program then type password" << std::endl;
        std::cout << "else if you consider ot exit then type 'exit' ?????  ";
        std::cout << "                                                  ??????? ";
        std::cin >> msg;
        std::transform(msg.begin(), msg.end(), msg.begin(), [](unsigned char c) {return std::tolower(c); });
        if (msg == "exit") {
            return false;
        }
    }
    return true;
}

int main(int argc, char const* argv[]) {

    std::string exePath = std::experimental::filesystem::path(argv[0]).parent_path().string();
    std::string givenFilePath = exePath + "/" + std::experimental::filesystem::path(argv[0]).stem().string() + ".xlsx";
    // std::wstring_convert<std::codecvt_utf8<wchar_t>> converter;
    // std::wstring wideString = converter.from_bytes(givenFilePath);
    // const wchar_t* givenFile = wideString.c_str();


    if (welcomeMsg()) {
        if (!std::experimental::filesystem::exists(givenFilePath)) {
            std::cerr << "The prerequisite .xlsx file not found!!" << std::endl;
            return 0;
        }

        // StaffAllocation jay((std::wstring(givenFilePath.begin(),givenFilePath.end())).c_str());
        StaffAllocation jay(givenFilePath);
        try {
            jay.neo();
        }
        catch (std::exception& e) {
            std::cerr << "Exception caught: " << e.what() << std::endl;
            std::cout << "Press any key to exit...";
            std::cin.get();
        }
    }

    return 0;
}


/*

For Windows:

cmake -G "MinGW Makefiles"
mingw32-make

*/
