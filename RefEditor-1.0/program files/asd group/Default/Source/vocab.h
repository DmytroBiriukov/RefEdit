//---------------------------------------------------------------------------
#ifndef vocabH
#define vocabH
//---------------------------------------------------------------------------
#include <Classes.hpp>
#include <Controls.hpp>
#include <StdCtrls.hpp>
#include <Forms.hpp>
#include <Dialogs.hpp>
#include <utilcls.h>
#include "Word_2K_SRVR.h"
#include <OleServer.hpp>
struct  TRefCod2RefNo
{ long RefNo;
  wchar_t RefCode[5];
  unsigned int used;
};
void ref_sort(TRefCod2RefNo *a, int n);
long ref_search(wchar_t *key, TRefCod2RefNo *a, long i, long step);

//---------------------------------------------------------------------------
class TRefEditForm : public TForm
{
__published:	// IDE-managed Components
	TOpenDialog *OpenDialog1;
	TListBox *ListBox_Words;
        TButton *Button0;
	TWordApplication *WordApplication1;
	TWordDocument *WordDocument1;
        TButton *Button1;
        TButton *Button2;
        TListBox *ListBox1;
        TButton *Button3;
        TButton *Button4;
        TWordDocument *WordDocument2;

        void __fastcall Button_Click(TObject *Sender);
        void __fastcall FormCreate(TObject *Sender);
        void __fastcall Button1Click(TObject *Sender);
        void __fastcall Button2Click(TObject *Sender);
        void __fastcall Button3Click(TObject *Sender);
        void __fastcall Button4Click(TObject *Sender);
	
private:	// User declarations
	void __fastcall PrepareWord(void);
	long RefCode2RefNo(wchar_t* ref_code);
        void virtual __fastcall WMDropFiles(TWMDropFiles &message);
public:		// User declarations
	__fastcall TRefEditForm(TComponent* Owner);

	Variant	my_word;
	bool	can_run;
        int ref_count;
        AnsiString ref_fileName;
        TRefCod2RefNo* arr;
  BEGIN_MESSAGE_MAP
  MESSAGE_HANDLER(WM_DROPFILES, TWMDropFiles, WMDropFiles)
  END_MESSAGE_MAP(TForm);
};
//---------------------------------------------------------------------------
extern PACKAGE TRefEditForm *RefEditForm;
//---------------------------------------------------------------------------
#endif
