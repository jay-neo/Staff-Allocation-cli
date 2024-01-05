#include <iostream>
#include <string>
#include <vector>
#include <algorithm>
#include <random>
#include <filesystem>
#include "libxl.h"
#include <sstream>


std::string wchar2string(const char* charArray) {
    if (charArray==nullptr) {
        return "";
    }
    return std::string(charArray);
}

const char* string2wchar(const std::string& str) {
    return str.c_str();
}


bool validCell(std::string str) {
    for (char ch : str) {
        if (!std::isspace(ch)) {
            return true;
        }
    }
    return false;
}

int string2int(std::string str) {
    str.erase(std::remove_if(str.begin(), str.end(),
        [](char c) { return !std::isdigit(c); }), str.end());

    int res;
    std::istringstream(str) >> res;

    return res;
}

void Error(int sheetIdx, int l) {
    std::cout << std::endl;
    std::cout << "Line-" << l;
    std::cerr << " Something wrong in your Excel Sheet-" << sheetIdx + 1 << std::endl;
    std::cerr << "{May be number of staffs is less than the required number]" << std::endl;
    std::cout << "Press any key to exit..." << std::endl;
    std::cin.ignore(std::numeric_limits<std::streamsize>::max(), '\n');
    std::cin.get();
}

struct Staff {
    std::vector<std::pair<int, std::string>> Job;
    double total;
    int recent;
    int maxi;
    Staff() {
        total = 0;
        recent = 0;
        maxi = 1;
    }
    void addJob(int sft, std::string w) {
        total += (1.0 / static_cast<double>(maxi)) * 10.0;
        recent = 1;
        Job.emplace_back(sft, w);
    }
    void noJob() {
        recent = 0;
    }
};


struct Work {
    std::string name;
    int value;
    Work(std::string _name, int _value) {
        name = _name;
        value = _value;
    }
};


class StaffAllocation {
    int staffType = 0;
    int jobType = 0;
    libxl::Book* givenFile = nullptr;
    libxl::Book* staffFile = nullptr;
    std::string inputFileName;
    std::string outputFileName = "Final_Staff_Allocation_Sheet.xlsx";
    std::vector<std::pair<std::string, struct Staff>> staffs;
    std::vector<struct Work> currJobs;

public:
    StaffAllocation(std::string inputFilePath, int a, int b) : inputFileName(inputFilePath), staffType(a), jobType(b) {
        givenFile = xlCreateXMLBook();
        staffFile = xlCreateXMLBook();
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
        else if (s == "works") {
            std::shuffle(currJobs.begin(), currJobs.end(), g);
        }
    }

    void sortStaff() {
        std::sort(staffs.begin(), staffs.end(), [](const auto& lhs, const auto& rhs) {
            if (lhs.second.total != rhs.second.total) {
                return lhs.second.total < rhs.second.total;
            }
            else {
                return lhs.second.recent < rhs.second.recent;
            }
            });
    }


    void allocateSheet(const int& sheetIdx) {
        libxl::Sheet* givenSheet = givenFile->getSheet(sheetIdx);
        libxl::Sheet* staffSheet = staffFile->addSheet(givenSheet->name());

        int val = 1;
        if (jobType == 0) {
            std::cout << "Sheet " << sheetIdx + 1 << ": Enter the value of each work (number of staffs each room) ---> ";
            std::cin >> val;
        }

        int day = 0;
        int Rows = 1;
        int totalColumns = 1 + staffType;

        // Write Operation for Heading
        staffSheet->writeStr(0, 0, string2wchar("The code is open source in GitHub @jay-neo (https://github.com/jay-neo)"));
        staffSheet->writeStr(1, 0, string2wchar("Staff Name"));
        // staffSheet->writeStr(0, 0, givenSheet->readStr(0, 0));

        for (int c = 1 + staffType; c < givenSheet->lastCol(); c += (jobType + 1)) {
            std::string str = wchar2string(givenSheet->readStr(0, c));
            if (givenSheet->cellType(0, c) != libxl::CELLTYPE_BLANK and validCell(str)) {
                staffSheet->writeStr(1, day + 1, givenSheet->readStr(0, c));
                totalColumns += 2;
                ++day;
            }
        }

        // Storing staffs name into Staffs vector
        for (int r = 1; r < givenSheet->lastRow(); ++r) {
            std::string str = wchar2string(givenSheet->readStr(r, 0));
            if (givenSheet->cellType(r, 0) != libxl::CELLTYPE_BLANK and validCell(str)) {
                staffSheet->writeStr(r + 1, 0, givenSheet->readStr(r , 0));
                staffs.emplace_back(str, Staff());
                if (staffType) {
                    try {
                        std::string str2 = wchar2string(givenSheet->readStr(r, 1));
                        staffs[r - 1].second.maxi = string2int(str2);
                    } catch (const std::invalid_argument& e) {
                        std::cerr << "Invalid argument: " << e.what() << std::endl;
                        Error(sheetIdx, 1);
                        return;
                    }
                }
                ++Rows;
            }
        }


        int itr = 0;
        int req = 0;
        int workValue = jobType;
        std::vector<std::vector<struct Work>> totalJobs(day + 1);


        // Storing jobs into Jobs vector
        for (int c = (staffType + 1); c < totalColumns; c += (jobType + 1)) {

            req = 0;
            
            for (int r = 1; r < Rows; ++r) {
                std::string str1 = wchar2string(givenSheet->readStr(r, c));
                if (!validCell(str1)) {
                    continue;
                }
                if (jobType) {
                    try {
                        std::string str2 = wchar2string(givenSheet->readStr(r, c + 1));
                        workValue = string2int(str2);
                        req += workValue;
                    }
                    catch (const std::invalid_argument& e) {
                        std::cerr << "Invalid argument: " << e.what() << std::endl;
                        Error(sheetIdx, 2);
                        return;
                    }
                    totalJobs[itr].emplace_back(Work(str1, workValue));
                }
                else {
                    totalJobs[itr].emplace_back(Work(str1, val));
                }
            }

            if ((!jobType and (val * totalJobs[itr].size()) > Rows - 1) or (jobType and Rows - 1 < req)) {
                // std::cout << jobType << " " << totalJobs[itr].size() << " " << Rows << " " << req;
                Error(sheetIdx, 3);
                return;
            }

            ++itr;

        }

       
        this->shuffle("staffs");
        day = 0;

        // Doing allocation here
        for (int j = (staffType + 1); j < totalColumns; j += (jobType + 1)) {

            if (j != staffType + 1) {
                this->sortStaff();
            }

            currJobs = totalJobs[day];

            this->shuffle("works");

            itr = 0;
            for (int r = 0; r < currJobs.size(); ++r) {
                for (int i = 0; i < currJobs[r].value; ++i) {
                    if (staffs[itr].second.total >= 100 and staffType) {
                        Error(sheetIdx, 4);
                        std::cout << staffs[itr].first << " have to do extra job!!" << std::endl;
                    }
                    staffs[itr].second.addJob(day + 1, currJobs[r].name);
                    ++itr;
                }
            }

            while (itr < staffs.size()) {
                staffs[itr++].second.noJob();
            }

            currJobs.clear();
            ++day;
        }


        std::sort(staffs.begin(), staffs.end(), [](const auto& lhs, const auto& rhs) { return lhs.first < rhs.first; });

        // Final Write operation on Ouput Sheet
        for (itr = 1; itr < Rows; ++itr) {
            staffSheet->writeStr(itr + 1, 0, string2wchar(staffs[itr - 1].first));
            for (int i = 0; i < staffs[itr - 1].second.Job.size(); ++i) {
                staffSheet->writeStr(itr + 1, staffs[itr - 1].second.Job[i].first , string2wchar(staffs[itr - 1].second.Job[i].second));
            }
        }

        staffs.clear();


        std::cout << std::endl << "Sheet-" << sheetIdx + 1 << ": Staff Allocation Complete." << std::endl;

    }

    void neo() {
        if (givenFile->load(string2wchar(inputFileName))) {

            for (int sheetIdx = 0; sheetIdx < givenFile->sheetCount(); ++sheetIdx) {
                this->allocateSheet(sheetIdx);
            }
            staffFile->save(string2wchar(outputFileName));

        }
        else {
            std::cerr << "Input Excel File(.xlsx) not found !!" << std::endl;
            std::cout << "Press any key to exit...";
            std::cin.get();
        }

        std::cout << std::endl << std::endl;
        std::cout << "               ####     ####   #    #  ####### " << std::endl;
        std::cout << "               #   #   #    #  ##   #  #       " << std::endl;
        std::cout << "               #    #  #    #  # #  #  #####   " << std::endl;
        std::cout << "               #    #  #    #  #  # #  #       " << std::endl;
        std::cout << "               #   #   #    #  #   ##  #       " << std::endl;
        std::cout << "               ####     ####   #    #  ######  " << std::endl;
        std::cout << std::endl << std::endl;
        std::cin.ignore(std::numeric_limits<std::streamsize>::max(), '\n');

    }
};



bool welcomeMsg(int &a, int &b) {

    std::string msg = "0";
    std::cout << std::endl;
    std::cout << "      #################################################" << std::endl;
    std::cout << "      #      Enter the submitted staff's type         #" << std::endl;
    std::cout << "      #      Type 1 : Staffs with unlimited worklife  #" << std::endl;
    std::cout << "      #      Type 2 : Staffs with limited worklife    #" << std::endl;
    std::cout << "      #      Type 0 : Exit from the program           #" << std::endl;
    std::cout << "      #################################################" << std::endl;
    std::cout << "      Enter your choice ---------> ";
    while (msg != "1" or msg != "2") {
        std::cin >> msg;
        if (msg == "1") {
            a = 0; break;
        }
        else if (msg == "2") {
            a = 1; break;
        }
        else if (msg == "0") {
            return false;
        }
        std::cout << "Enter your valid choice ---> ";
    }
    std::cout << std::endl << std::endl;


    std::cout << "      ########################################################" << std::endl;
    std::cout << "      #      Enter the submitted job's type                  #" << std::endl;
    std::cout << "      #      Type 1 : Each job with equal requirement        #" << std::endl;
    std::cout << "      #      Type 2 : Each job with different requirement    #" << std::endl;
    std::cout << "      #      Type 0 : Exit from the program                  #" << std::endl;
    std::cout << "      ########################################################" << std::endl;
    std::cout << "      Enter your choice ---------> ";
    while (msg != "1" or msg != "2") {
        std::cin >> msg;
        if (msg == "1") {
            b = 0; break;
        }
        else if (msg == "2") {
            b = 1; break;
        }
        else if (msg == "0") {
            return false;
        }
        std::cout << "Enter your valid choice ---> ";
    }
    std::cout << std::endl << std::endl;


    int j = 0;
    std::cout << "Prerequisites:" << std::endl << std::endl;
    std::cout << "1. Excel file(.xlsx) should be formatted like this:" << std::endl;
    std::cout << "###############";
    if (a) {
        std::cout << "###############";
    }
    for (int j = 1; j <= 3; ++j) {
        std::cout << "###############";
        if (b) {
            std::cout << "###############";
        }
    }
    std::cout << std::endl;
    std::cout << "#   Staff Name  #";
    if (a) {
        std::cout << "  WorkLife   #";
    }
    for (int j = 1; j <= 3 ; ++j) {
        std::cout << "     Day " << j << "   #";
        if (b) {
            std::cout << " Requirement  #";
        }
    }
    std::cout << std::endl;
    int itr = 1;
    for (int i = 0; i <= 14; ++i) {
        if (i % 2 == 0) {
            std::cout << "###############";
            if (a) {
                std::cout << "###############";
            }
            for (int j = 1; j <= 3; ++j) {
                std::cout << "###############";
                if (b) {
                    std::cout << "###############";
                }
            }
        }
        else {
            std::cout << "#   Person " << itr << "    #";
            if (a) {
                std::cout << "  Integer " << itr << "  #";
            }
            for (int j = 1; j <= 3; ++j) {
                std::cout << "     Job " << itr << "   #";
                if (b) {
                    std::cout << " Integer " << itr << "    #";
                }
            }
            ++itr;
        }
        std::cout << std::endl;
    }
    std::cout << std::endl;
    std::cout << "2. The excel file name should be same name as the execution file(.exe)" << std::endl << std::endl;

    while (msg != "jay-neo") {
        std::cout << "If you understand the prerequisites and continue the program then type password" << std::endl;
        std::cout << "else if you consider to exit then type 'exit' -------> ";
        std::cin >> msg;
        std::transform(msg.begin(), msg.end(), msg.begin(), [](unsigned char c) {return std::tolower(c); });
        if (msg == "exit") {
            return false;
        }
    }
    return true;
}

int main(int argc, char const* argv[]) {

    std::string exePath = std::filesystem::path(argv[0]).parent_path().string();
    std::string givenFilePath = exePath + "/" + std::filesystem::path(argv[0]).stem().string() + ".xlsx";

    int a, b;

    if (welcomeMsg(a, b)) {
        if (!std::filesystem::exists(givenFilePath)) {
            std::cerr << "The prerequisite .xlsx file not found!!" << std::endl;
            return 0;
        }

        StaffAllocation jay(givenFilePath, a , b);
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


