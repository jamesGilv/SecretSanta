// SecretSanta.cpp : This file contains the 'main' function. Program execution begins and ends there.

#include<tchar.h>
#include<Windows.h>
#include<vector>
#include<random>
#include<iostream>
#include<iomanip>
#include<fstream>
#include<string>
#include<locale>
#include<codecvt>

#include "EASendMailObj.tlh"
using namespace EASendMailObjLib;

const int ConnectNormal = 0;
const int ConnectSSLAuto = 1;
const int ConnectSTARTTLS = 2;
const int ConnectDirectSSL = 3;
const int ConnectTryTLS = 4;

std::vector<std::string> shift(std::vector<std::string>& in)
{
    int length = in.size();
    std::random_device dev;
    std::mt19937 rng(dev());
    std::uniform_int_distribution<std::mt19937::result_type> dist(1, length-1);
    std::vector<std::string> dummy;
    // create random shift from 1 to length of names
    int move = dist(rng);
    for (auto it = move; it != length; ++it)
    {
        // create dummy vector from shifted position to end
        dummy.push_back(in[it]);
    }
    for (auto it = 0; it != move; ++it)
    {
        // go from start to shifted position
        dummy.push_back(in[it]);
    }
    return dummy;
}

// Function to convert std::string to std::wstring
std::wstring stringToWString(const std::string& str) 
{
    std::wstring_convert<std::codecvt_utf8<wchar_t>, wchar_t> converter;
    return converter.from_bytes(str);
}

void send_mail(std::string& recip, std::string name)
{
    std::wstring wideRecip = stringToWString(recip);
    std::string message = "This year, you will be buying presents for " + name;
    // Convert std::string to std::wstring
    std::wstring wideMessage = stringToWString(message);
    ::CoInitialize(NULL);
    try
    {
        IMailPtr oSmtp = NULL;
        oSmtp.CreateInstance(__uuidof(EASendMailObjLib::Mail));
        oSmtp->LicenseCode = _T("TryIt");

        // Set your hotmail/outlook/office 365 email address
        oSmtp->FromAddr = _T("YOUREMAIL@DOMAIN.COM");
        // Add recipient using the wide string
        oSmtp->AddRecipientEx(wideRecip.c_str(), 0);
        // Set email subject
        oSmtp->Subject = _T("Secret Santa");
        // Assign the wide string to BodyText
        oSmtp->BodyText = wideMessage.c_str();

        // Hotmail/Outlook/Office365 SMTP server address
        oSmtp->ServerAddr = _T("smtp.office365.com");
        // Hotmail user authentication should use your
        // Hotmail email address as the user name.
        oSmtp->UserName = _T("YOUREMAIL@DOMAIN.com");
        // Password
        oSmtp->Password = _T("YOURPASSWORD");
        // set 587 port
        oSmtp->ServerPort = 587;
        // detect SSL/TLS automatically.
        oSmtp->ConnectType = ConnectSSLAuto;

        // oSmtp->SendMail sends the email, need to have a single recipient for each email
        oSmtp->SendMail();
        std::wcout << L"Email sent to " << wideRecip << std::endl;
    }
    catch(const _com_error& e)
    {
        std::wcerr << L"Error sending email: " << e.ErrorMessage() << std::endl;
    }
    CoUninitialize();
}

int _tmain(int argc, _TCHAR* argv[])
{
    // Define variables
    std::string data_file{ "People.txt" };
    // Expected format of People.txt is
    // FirstName LastName FirstLast@domain.com
    std::vector<std::string> data;
    std::vector<std::string> names;
    std::vector<std::string> emails;

    // Open file (you must check if successful)
    std::fstream name_stream(data_file);
    if (!name_stream)
    {
        std::cerr << "File could not be found";
        return 1;
    }
    while (name_stream)
    {
        // copy each line to vector
        std::string dummy_string;
        std::getline(name_stream, dummy_string);
        if (std::all_of(dummy_string.begin(), dummy_string.end(), isspace))
            continue;
        data.push_back(dummy_string);
    }
    // Close file after reading to vector
    name_stream.close();
    std::cout << "File read successfully" << '\n';

    // shuffle order of data randomly
    std::random_shuffle(data.begin(), data.end());

    // Create a temporary string to hold each word
    std::string tempName;
    std::string tempMail;

    // Iterate over each character in the sentence
    for (std::vector<std::string>::iterator i{ data.begin() }; i != data.end(); i++)
    {
        char keep{ 'n' };
        // find where name ends and email begins
        std::size_t found = i->rfind(" ");
        tempName = i->substr(0, found);
        tempMail = i->substr(found + 1);
        // ask if person is in secret santa
        std::cout << "Is " << tempName << " participating in secret santa this year y/n\n";
        std::cin >> keep;
        if (keep == 'y' || keep == 'Y')
        {
            // copy name and email to separate vectors
            names.push_back(tempName);
            emails.push_back(tempMail);
        }
    }

    names = shift(names);

    for (int i{ 0 }; i < names.size(); i++)
    {
        std::string name = names[i];
        std::string email = emails[i];
        send_mail(email, name);
    }
    
    return 0;
}
