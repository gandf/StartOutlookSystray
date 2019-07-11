//---------------------------------------------------------------------------

#include <vcl.h>
#pragma hdrstop

#include "Unit1.h"
//---------------------------------------------------------------------------
#pragma package(smart_init)
#pragma resource "*.dfm"
TForm1 *Form1;
//---------------------------------------------------------------------------
__fastcall TForm1::TForm1(TComponent* Owner)
    : TForm(Owner)
{
}
//---------------------------------------------------------------------------
void __fastcall TForm1::LangChange(TObject *Sender)
{
switch (Lang->ItemIndex)
       {
       case 0: {
               //English
               OutlookParameterLbl->Caption = "Additionnal parameters:";
               OutlookPathLbl->Caption = "Path to Outlook:";
               RunOnStartup->Caption = "&Run on Windows Startup";
               DeleteFromStartup->Caption = "&Delete from Windows Startup";
               MethodLbl->Caption = "Method:";
               WaitLbl->Caption = "Wait to minimize Outlook:";
               WaitBeforeLbl->Caption = "Wait before run Outlook:";
               WaitMsBeforeLbl->Caption = "ms";
               WaitMsLbl->Caption = "ms";
               Test->Caption = "&Test";
               Save->Caption = "&Save";
               UpdateFromRegistry->Caption = "Update Path from &Registry";
               TextOutlookRunning = "Outlook is already running.";
               break;
               }
       case 1: {
               //Français
               OutlookParameterLbl->Caption = "Paramètres supplémentaires :";
               OutlookPathLbl->Caption = "Chemin vers Outlook :";
               RunOnStartup->Caption = "&Lancer au démarrage de Windows";
               DeleteFromStartup->Caption = "&Supprimer du démarrage de Windows";
               MethodLbl->Caption = "Méthode :";
               WaitLbl->Caption = "Attente pour minimiser Outlook :";
               WaitBeforeLbl->Caption = "Attente avant de lancer Outlook :";
               WaitMsBeforeLbl->Caption = "ms";
               WaitMsLbl->Caption = "ms";
               Test->Caption = "&Tester";
               Save->Caption = "&Sauvegarder";
               UpdateFromRegistry->Caption = "Mise à jour depuis la base de &Registre";
               TextOutlookRunning = "Outlook est déja en cours d'exécution.";
               break;
               }
       }
CheckRunOnStartup();
}
//---------------------------------------------------------------------------
void __fastcall TForm1::RunOnStartupClick(TObject *Sender)
{
if (Method->ItemIndex < 0)
   {
   Method->ItemIndex = 0;
   }
if (WaitBefore->Text.ToIntDef(0) < 0)
   {
   WaitBefore->Text = 0;
   }
if (WaitEdit->Text.ToIntDef(1500) < 0)
   {
   WaitEdit->Text = 1500;
   }
TRegistry *Reg = NULL;
try {
    UnicodeString Param = "\"" + IntToStr(WaitBefore->Text.ToIntDef(0)) + "\" \"" + IntToStr(WaitEdit->Text.ToIntDef(1500)) + "\" \"" + IntToStr(Method->ItemIndex) + "\"";
    Reg = new TRegistry();
    Reg->RootKey = HKEY_CURRENT_USER;
    Reg->OpenKey("Software\\Microsoft\\Windows\\CurrentVersion\\Run", false);
    if (Reg->ValueExists("OutlookAutoSystray"))
      {
      Reg->DeleteValue("OutlookAutoSystray");
      Reg->WriteString("OutlookAutoSystray", "\"" + Application->ExeName + "\" \"" + OutlookPath->Text + "\" " + Param + " \"" + OutlookParameter->Text + "\"");
      }
    else {
         Reg->WriteString("OutlookAutoSystray", "\"" + Application->ExeName + "\" \"" + OutlookPath->Text + "\" " + Param + " \"" + OutlookParameter->Text + "\"");
         }
    Reg->CloseKey();
    }
catch (...)
      {
      if (Reg != NULL)
         {
         try {
             delete Reg;
             }
         catch (...)
               {
               }
         Reg = NULL;
         }
      }
CheckRunOnStartup();
}
//---------------------------------------------------------------------------
UnicodeString __fastcall TForm1::ReadRegistry(bool FirstTry)
{
TRegistry *Reg = NULL;
UnicodeString CLSID = "";
UnicodeString Result = "";
try {
    #if _WIN64
    // 64-bit Windows
    if (FirstTry)
       {
       Reg = new TRegistry(KEY_READ | KEY_WOW64_64KEY);
       }
    else {
         Reg = new TRegistry(KEY_READ | KEY_WOW64_32KEY);
         }
    #elif _WIN32
    // 32-bit Windows
    int Is64;
    if (!IsWow64Process(GetCurrentProcess(), &Is64))
       {
       Is64 = false;
       }
    if (Is64)
       {
       //app 32-bit on Windows 64
       if (FirstTry)
          {
          Reg = new TRegistry(KEY_READ | KEY_WOW64_64KEY);
          }
       else {
            Reg = new TRegistry(KEY_READ | KEY_WOW64_32KEY);
            }
       }
    else {
         //app 32-bit on Windows 32
         if (!FirstTry)
            {
            return Result;
            }
         //Search Outlook 32 bit
         Reg = new TRegistry(KEY_READ);
         }
    #endif
    Reg->RootKey = HKEY_LOCAL_MACHINE;
    if (Reg->OpenKey("Software\\Classes\\Outlook.Application\\CLSID", false))
       {
       if (Reg->ValueExists(""))
          {
          CLSID = Reg->ReadString("");
          Reg->CloseKey();
          Reg->RootKey = HKEY_LOCAL_MACHINE;
          if (Reg->OpenKey("Software\\Classes\\CLSID\\" + CLSID + "\\LocalServer32", false))
             {
             if (Reg->ValueExists(""))
                {
                Result = Reg->ReadString("");
                }
             }
          }
       }
    Reg->CloseKey();
    delete Reg;
    Reg = NULL;
    }
catch (...)
      {
      if (Reg != NULL)
         {
         try {
             Reg->CloseKey();
             delete Reg;
             }
         catch (...)
               {
               }
         Reg = NULL;
         }
      }
return Result;
}
//---------------------------------------------------------------------------
void __fastcall TForm1::FormCreate(TObject *Sender)
{
if (ParamCount() > 0)
   {
   if (processExists(ExtractFileName(ParamStr(1))))
      {
      Application->ShowMainForm = false;
      Application->Terminate();
      }
   //Vérifie si une autre instance est en cours
   UnicodeString MagicFile = TPath::GetTempPath() + ChangeFileExt(ExtractFileName(Application->ExeName), ".check");

   Wait(10, 1500);

   bool CanRun = true;
   if (FileExists(MagicFile))
      {
      TDateTime FileDateTime;
      TTime ToDeduct = EncodeTime(0, 1, 0, 0);
      if (FileAge(MagicFile, FileDateTime))
         {
         if (FileDateTime < (FileDateTime.CurrentDateTime() - ToDeduct))
            {
            DeleteFile(MagicFile);
            }
         else {
              CanRun = false;
              }
         }
      else {
           CanRun = false;
           }
      }
   if (CanRun)
      {
      //Creation du fichier temporaire
      TFileStream* FileStr = NULL;
      try {
          FileStr = TFile::Create(MagicFile);
          STARTUPINFO StartInfo;
          PROCESS_INFORMATION ProcInfo;
          memset(&ProcInfo, 0, sizeof(ProcInfo));
          memset(&StartInfo, 0 , sizeof(StartInfo));
          StartInfo.cb = sizeof(StartInfo);
          StartInfo.dwFlags = STARTF_USESHOWWINDOW;
          StartInfo.wShowWindow = SW_MAXIMIZE;
          if (processExists(ExtractFileName(ParamStr(1))))
             {
             Application->ShowMainForm = false;
             Application->Terminate();
             }
          Sleep(ParamStr(2).ToIntDef(0));
          if (CreateProcess(NULL,("\"" + ParamStr(1) + "\" " + ParamStr(5)).c_str(), NULL, NULL, 0, 0, NULL, NULL, &StartInfo, &ProcInfo))
             {
             WaitForInputIdle(ProcInfo.hProcess, INFINITE);
             Sleep(ParamStr(3).ToIntDef(1500));
             EnumWindows((WNDENUMPROC)EnumWindowsProc, (LPARAM)ParamStr(4).ToIntDef(0));
             Wait(2000, 8000);
             }
          }
      catch (...)
            {
            }
      try {
          if (FileStr != NULL)
             {
             delete FileStr;
             }
          }
      catch (...)
            {
            }
      try {
          if (FileExists(MagicFile))
             {
             DeleteFile(MagicFile);
             }
          }
      catch (...)
            {
            }
      }
   Application->ShowMainForm = false;
   Application->Terminate();
   }

UpdatePath();

if (FileExists(ChangeFileExt(Application->ExeName,".ini")))
   {
   TIniFile *ini = NULL;
   try {
       ini = new TIniFile(ChangeFileExt(Application->ExeName,".ini"));
       WaitBefore->Text = IntToStr(ini->ReadInteger("Setup", "WaitBefore", 0));
       OutlookPath->Text = ini->ReadString("Setup", "OutlookPath", OutlookPath->Text);
       OutlookParameter->Text = ini->ReadString("Setup", "OutlookParameter", "/recycle");
       Lang->ItemIndex = ini->ReadInteger("Setup", "Lang", 0);
       Method->ItemIndex = ini->ReadInteger("Setup", "Method", 0);
       WaitEdit->Text = IntToStr(ini->ReadInteger("Setup", "Wait", 1500));

       Form1->Top = ini->ReadInteger("Setup", "Top", 100);
       Form1->Left = ini->ReadInteger("Setup", "Left", 250);
       if ((Form1->Top < 0) || (Form1->Top > Screen->DesktopHeight))
          {
          Form1->Top = 0;
          }
       if ((Form1->Left < 0) || (Form1->Left > Screen->DesktopWidth))
          {
          Form1->Left = 0;
          }
       }
   catch (...)
         {
         if (ini != NULL)
            {
            try {
                delete ini;
                }
            catch (...)
                  {
                  }
            ini = NULL;
            }
         }
   }
}
//---------------------------------------------------------------------------
void __fastcall TForm1::TestClick(TObject *Sender)
{
if (processExists(ExtractFileName(OutlookPath->Text)))
   {
   Application->MessageBox(TextOutlookRunning.c_str(), Application->Title.c_str(), MB_OK + MB_ICONINFORMATION + MB_TOPMOST);
   return;
   }

try {
    STARTUPINFO StartInfo;
    PROCESS_INFORMATION ProcInfo;
    memset(&ProcInfo, 0, sizeof(ProcInfo));
    memset(&StartInfo, 0 , sizeof(StartInfo));
    StartInfo.cb = sizeof(StartInfo);
    StartInfo.dwFlags = STARTF_USESHOWWINDOW;
    StartInfo.wShowWindow = SW_MAXIMIZE;
    Sleep(WaitBefore->Text.ToIntDef(0));
    if (CreateProcess(NULL,("\"" + OutlookPath->Text + "\" " + OutlookParameter->Text).c_str(), NULL, NULL, 0, 0, NULL, NULL, &StartInfo, &ProcInfo))
       {
       WaitForInputIdle(ProcInfo.hProcess, INFINITE);
       Sleep(WaitEdit->Text.ToIntDef(1500));
       EnumWindows((WNDENUMPROC)EnumWindowsProc, (LPARAM)Method->ItemIndex);
       }

    }
catch ( const Exception &Exp)
      {
      Application->MessageBox(Exp.Message.c_str(), Application->Title.c_str(), MB_OK + MB_ICONINFORMATION + MB_TOPMOST);
      }
}
//---------------------------------------------------------------------------
bool CALLBACK EnumWindowsProc(HWND hWnd, LPARAM lParam)
{
try {
    wchar_t WindowName[80], ClassName[80];
    GetWindowText(hWnd, WindowName, 80);
    GetClassName(hWnd, ClassName, 80);
    if (UnicodeString(ClassName) == "rctrl_renwnd32")
       {
       HWND Win = FindWindow(ClassName, WindowName);
       switch ((int)lParam)
              {
              case 0: {
                      PostMessage(Win, WM_SYSCOMMAND, SC_MINIMIZE, 0);
                      break;
                      }
              case 1: {
                      ShowWindow(Win, SW_MINIMIZE);
                      break;
                      }
              }
       }
    }
catch (...)
      {
      }
return true;
}
//---------------------------------------------------------------------------
void __fastcall TForm1::FormKeyPress(TObject *Sender, System::WideChar &Key)
{
if (Key == vkEscape)
   {
   Form1->Close();
   }
}
//---------------------------------------------------------------------------
void __fastcall TForm1::SaveClick(TObject *Sender)
{
TIniFile *ini = NULL;
try {
    ini = new TIniFile(ChangeFileExt(Application->ExeName,".ini"));
    ini->WriteInteger("Setup", "WaitBefore", WaitBefore->Text.ToIntDef(0));
    ini->WriteString("Setup", "OutlookPath", OutlookPath->Text);
    ini->WriteString("Setup", "OutlookParameter", OutlookParameter->Text);
    ini->WriteInteger("Setup", "Lang", Lang->ItemIndex);
    ini->WriteInteger("Setup", "Method", Method->ItemIndex);
    ini->WriteInteger("Setup", "Wait", WaitEdit->Text.ToIntDef(1500));
    ini->WriteInteger("Setup", "Top", Form1->Top);
    ini->WriteInteger("Setup", "Left", Form1->Left);
    }
catch (...)
      {
      if (ini != NULL)
         {
         try {
             delete ini;
             }
         catch (...)
               {
               }
         ini = NULL;
         }
      }
}
//---------------------------------------------------------------------------
void __fastcall TForm1::FormMouseDown(TObject *Sender, TMouseButton Button, TShiftState Shift, int X, int Y)
{
ReleaseCapture();
#if _WIN64
  // 64-bit Windows
  Perform(WM_SYSCOMMAND, 0xF012, (NativeInt)0);
#elif _WIN32
  // 32-bit Windows
  Perform(WM_SYSCOMMAND, 0xF012, 0);
#endif
}
//---------------------------------------------------------------------------
void __fastcall TForm1::WaitEditChange(TObject *Sender)
{
TEdit* TempEdit = (TEdit *)Sender;
int ValDef = 0;
if (TempEdit->Name == WaitEdit->Name)
   {
   ValDef = 1500;
   }
if (TempEdit->Text.ToIntDef(ValDef) >= 0)
   {
   TempEdit->Text = IntToStr(TempEdit->Text.ToIntDef(ValDef));
   }
else {
     TempEdit->Text = 0;
     }
}
//---------------------------------------------------------------------------
void __fastcall TForm1::DeleteFromStartupClick(TObject *Sender)
{
TRegistry *Reg = NULL;
try {
    Reg = new TRegistry();
    Reg->RootKey = HKEY_CURRENT_USER;
    Reg->OpenKey("Software\\Microsoft\\Windows\\CurrentVersion\\Run", false);
    if (Reg->ValueExists("OutlookAutoSystray"))
      {
      Reg->DeleteValue("OutlookAutoSystray");
      }
    Reg->CloseKey();
    }
catch (...)
      {
      if (Reg != NULL)
         {
         try {
             delete Reg;
             }
         catch (...)
               {
               }
         Reg = NULL;
         }
      }
CheckRunOnStartup();
}
//---------------------------------------------------------------------------
void __fastcall TForm1::CheckRunOnStartup()
{
DeleteFromStartup->Enabled = false;
TRegistry *Reg = NULL;
try {
    Reg = new TRegistry();
    Reg->RootKey = HKEY_CURRENT_USER;
    Reg->OpenKey("Software\\Microsoft\\Windows\\CurrentVersion\\Run", false);
    if (Reg->ValueExists("OutlookAutoSystray"))
      {
      DeleteFromStartup->Enabled = true;
      }
    Reg->CloseKey();
    }
catch (...)
      {
      if (Reg != NULL)
         {
         try {
             delete Reg;
             }
         catch (...)
               {
               }
         Reg = NULL;
         }
      }
}
//---------------------------------------------------------------------------
void __fastcall TForm1::UpdateFromRegistryClick(TObject *Sender)
{
UpdatePath();
}
//---------------------------------------------------------------------------
void __fastcall TForm1::UpdatePath()
{
UnicodeString OldOutlookPath = OutlookPath->Text;
OutlookPath->Text = ReadRegistry(true);
if (OutlookPath->Text == "")
   {
   OutlookPath->Text = ReadRegistry(false);
   }

if (OutlookPath->Text == "")
   {
   OutlookPath->Text = OldOutlookPath;
   }
}
//---------------------------------------------------------------------------
bool __fastcall TForm1::processExists(UnicodeString ExeFileName)
{
bool Result = false;
try {
    PROCESSENTRY32 FProcessEntry32;
    FProcessEntry32.dwSize = sizeof(PROCESSENTRY32);
    HANDLE FSnapshot = CreateToolhelp32Snapshot(TH32CS_SNAPPROCESS, NULL);
    if (Process32First(FSnapshot, &FProcessEntry32))
       {
       do {
          if (!_wcsicmp(FProcessEntry32.szExeFile, ExeFileName.c_str()))
             {
             Result = true;
             break;
             }
          }
       while (Process32Next(FSnapshot, &FProcessEntry32));
       }

    CloseHandle(FSnapshot);
    }
catch (...)
      {
      }
return Result;
}
//---------------------------------------------------------------------------
void __fastcall TForm1::Wait(int Min, int Max)
{
#if _WIN64
std::mt19937_64 eng{std::random_device{}()};  // or seed however you want
std::uniform_int_distribution<> dist{Min, Max};
std::this_thread::sleep_for(std::chrono::milliseconds{dist(eng)});
#elif _WIN32
srand(time(NULL));
Sleep((std::rand()% (Max - Min + 1)) + Min);
#endif
}
//---------------------------------------------------------------------------

