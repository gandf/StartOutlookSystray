//---------------------------------------------------------------------------

#ifndef Unit1H
#define Unit1H
//---------------------------------------------------------------------------
#include <System.Classes.hpp>
#include <Vcl.Controls.hpp>
#include <Vcl.StdCtrls.hpp>
#include <Vcl.Forms.hpp>
#include <System.Win.Registry.hpp>
#include <Windows.h>
#include <Vcl.ComCtrls.hpp>
#include "cspin.h"
#include <IniFiles.hpp>
//---------------------------------------------------------------------------
bool CALLBACK EnumWindowsProc(HWND hWnd, LPARAM lParam);
class TForm1 : public TForm
{
__published:	// Composants gérés par l'EDI
    TEdit *OutlookPath;
    TLabel *OutlookPathLbl;
    TEdit *OutlookParameter;
    TLabel *OutlookParameterLbl;
    TComboBox *Lang;
    TButton *RunOnStartup;
    TButton *Test;
    TComboBox *Method;
    TLabel *MethodLbl;
    TLabel *WaitLbl;
    TButton *Save;
    TEdit *WaitEdit;
    TLabel *WaitMsLbl;
    TLabel *WaitBeforeLbl;
    TEdit *WaitBefore;
    TLabel *WaitMsBeforeLbl;
    TButton *DeleteFromStartup;
    void __fastcall LangChange(TObject *Sender);
    void __fastcall RunOnStartupClick(TObject *Sender);
    void __fastcall FormCreate(TObject *Sender);
    void __fastcall TestClick(TObject *Sender);
    void __fastcall FormKeyPress(TObject *Sender, System::WideChar &Key);
    void __fastcall SaveClick(TObject *Sender);
    void __fastcall FormMouseDown(TObject *Sender, TMouseButton Button, TShiftState Shift,
          int X, int Y);
    void __fastcall WaitEditChange(TObject *Sender);
    void __fastcall DeleteFromStartupClick(TObject *Sender);
private:	// Déclarations utilisateur
    void __fastcall CheckRunOnStartup();
public:		// Déclarations utilisateur
    __fastcall TForm1(TComponent* Owner);
};
//---------------------------------------------------------------------------
extern PACKAGE TForm1 *Form1;
//---------------------------------------------------------------------------
#endif
