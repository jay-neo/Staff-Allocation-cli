#include <iostream>
#include <cstring>
#include <vector>
#include <algorithm>
#include <random>
#include <filesystem>
#include "libxl.h"


struct Staff {
    std::vector<std::pair<int, const char*>> Job;
    int total;
    int recent;
    Staff() {
        total = 0;
        recent = 0;
    }
    void addJob(int sft, const char* w){
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
    libxl::Book* givenFile = xlCreateXMLBook();
    libxl::Book* staffFile = xlCreateXMLBook();
    const char* inputFileName;
    const char* outputFileName = "Staff Allocation Sheet.xlsx";
    std::vector<std::pair<const char*, struct Staff*>> staffs;
    std::vector<const char*> rooms;

public:
    StaffAllocation(const char* inputFilePath, int _st = 2) : inputFileName(inputFilePath), st(_st) {
        // givenFile = xlCreateXMLBook();
        // staffFile = xlCreateXMLBook();   
    }

    ~StaffAllocation() {
        givenFile->release();
        staffFile->release();
    }
    
    void shuffle(const std::string &s) {
        std::random_device rd;
        std::mt19937 g(rd());
        if (s=="staffs") {
            std::shuffle(staffs.begin(), staffs.end(), g);
        } else if (s=="rooms") {
            std::shuffle(rooms.begin(), rooms.end(), g);
        }
    }

    void sortStaff() {
        std::sort(staffs.begin(), staffs.end(), [](const auto &lhs, const auto &rhs){
            if (lhs.second->total != rhs.second->total){ // error
                return lhs.second->total < rhs.second->total;
            } else if (lhs.second->recent != rhs.second->recent){
                return lhs.second->recent < rhs.second->recent;
            } else {
                return std::strcmp(lhs.first, rhs.first) < 0;
            }
        });
    }

    
    void allocateSheet(const int &sheetIdx) {
        libxl::Sheet* givenSheet = givenFile->getSheet(sheetIdx);
        libxl::Sheet* staffSheet = staffFile->addSheet(givenSheet->name());

        int Rows = givenSheet->lastRow(); // +1
        int Columns = givenSheet->lastCol(); //+1
        int itr = -1;

        staffSheet->writeStr(0, 0, givenSheet->readStr(0, 0));

        for(int r=1; r<Rows; ++r) {
            if (givenSheet->cellType(r, 0) != libxl::CELLTYPE_BLANK){
                staffs.emplace_back(givenSheet->readStr(r, 0), new Staff());
                staffSheet->writeStr(r, 0, givenSheet->readStr(r, 0));
                ++itr;
            }
        }
        int totalStaff = itr;

        this->shuffle("staffs");

        for(int shift=1; shift<Columns; ++shift) {
            // after first shift sort the 
            if(shift!=1){
                this->sortStaff();
            }

            // calculate room number for each shift
            int rows = 1;
            for(int r=1; r<Rows; ++r) {
                if(givenSheet->cellType(r, shift) != libxl::CELLTYPE_BLANK) {
                    ++rows;
                    rooms.emplace_back(givenSheet->readStr(r, shift));
                }
            }

            this->shuffle("rooms");
            staffSheet->writeStr(0, shift, givenSheet->readStr(0, shift));

            // Final allocate for each shift
            itr = 0;
            for(int r=0; r<rows; ++r) {
                for(int i=0; i<st; ++i){
                    staffs[itr++].second->addJob(shift, rooms[r]);
                }    
            }

            // those not hired, those (recent = 0)
            while(itr<=totalStaff){
                staffs[itr++].second->addJob();
            }

        }
        for(itr = 0; itr<totalStaff; ++itr) {
            for(int i=0; i<staffs[itr].second->Job.size(); ++i) {
                staffSheet->writeStr(itr + 1, staffs[itr].second->Job[i].first, staffs[itr].second->Job[i].second);
            }
        }

    }

    void neo() {
        if(givenFile->load(inputFileName)){
            int numSheet = givenFile->sheetCount();

            for(int sheetIdx=0; sheetIdx<numSheet; ++sheetIdx){
                this->allocateSheet(sheetIdx);
            }
            staffFile->save(outputFileName);

        } else {
            std::cerr << "Input Excel File not found !!" << std::endl;
            std::cout << "Press any key to exit...";
            std::cin.get();
        }
    }
};


bool welcomeMsg() {

    std::string msg = "no";
    std::cout << "Prerequisites:" << std::endl;
    std::cout << "1. Excel file(.xlsx) should be formatted like this 👇" << std::endl;
    std::cout << "╔═════════════╦═════════╦═════════╦═════════╦═════════╗" << std::endl;
    std::cout << "║ Staff Name  ║ Shift 1 ║ Shift 2 ║ Shift 3 ║ Shift n ║" << std::endl;
    std::cout << "╠═════════════╬═════════╬═════════╬═════════╬═════════║" << std::endl;
    std::cout << "║ Jagadish S  ║   A100  ║   A101  ║   A100  ║   B100  ║" << std::endl;
    std::cout << "╠═════════════╬═════════╬═════════╬═════════╬═════════║" << std::endl;
    std::cout << "║  Subarna M  ║   A101  ║   B100  ║         ║   B101  ║" << std::endl;
    std::cout << "╠═════════════╬═════════╬═════════╬═════════╬═════════║" << std::endl;
    std::cout << "║ Soumyajit N ║         ║   B101  ║         ║         ║" << std::endl;
    std::cout << "╠═════════════╬═════════╬═════════╬═════════╬═════════║" << std::endl;
    std::cout << "║  Sourita C  ║         ║         ║         ║         ║" << std::endl;
    std::cout << "╠═════════════╬═════════╬═════════╬═════════╬═════════║" << std::endl;
    std::cout << "║  person m   ║         ║         ║         ║         ║" << std::endl;
    std::cout << "╚═════════════╩═════════╩═════════╩═════════╩═════════╝" << std::endl;
    std::cout << "2. The excel file name should be same name as the execution file(.exe)" << std::endl;

    while(msg!="yes"){
        std::cout << "If you understand the prerequisites and continue the program then type 'yes'" << std::endl;
        std::cout << "else if you consider ot exit then type 'no' ┈┈┉┉┑  ";
        std::cout << "                                                └┄┄┄┄┄► ";
        std::cin >> msg;
        std::transform(msg.begin(), msg.end(), msg.begin(), [](unsigned char c){return std::tolower(c);});
        if(msg== "no"){
            return false;
        }
    }
    return true;
}

int main(int argc, char const *argv[]) {

    std::string exePath = std::filesystem::path(argv[0]).parent_path().string();
    std::string givenFilePath = exePath + "/" + std::filesystem::path(argv[0]).stem().string() + ".xlsx";
    // std::wstring_convert<std::codecvt_utf8<wchar_t>> converter;
    // std::wstring wideString = converter.from_bytes(givenFilePath);
    // const wchar_t* givenFile = wideString.c_str();

    
    if(welcomeMsg()){
        if (!std::filesystem::exists(givenFilePath)) {
            std::cerr << "The prerequisite .xlsx file not found!!" << std::endl;
            return 0;
        }
        
        StaffAllocation jay(givenFilePath.c_str());
        try {
            jay.neo();
        } catch (std::exception &e) {
            std::cerr << "Exception caught: " << e.what() << std::endl;
        }
    }

    return 0;
}

/*
For Linux:


*/

// cmake -G "MinGW Makefiles"
// mingw32-make