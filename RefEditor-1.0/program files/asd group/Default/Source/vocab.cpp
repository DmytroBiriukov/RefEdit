//---------------------------------------------------------------------------
#include <vcl.h>
#pragma hdrstop

#include "vocab.h"
//---------------------------------------------------------------------------
#pragma package(smart_init)
#pragma link "Word_2K_SRVR"
#pragma resource "*.dfm"
TRefEditForm *RefEditForm;
//---------------------------------------------------------------------------
void ref_sort(TRefCod2RefNo *a, int n)
{ TRefCod2RefNo tmp;
  for(int i=0; i<n; i++)
     for(int j=0; j<n; j++)
        if( wcscmp(a[i].RefCode,a[j].RefCode)<0)
          { tmp= a[i]; a[i]=a[j]; a[j]=tmp;
          }
}
//---------------------------------------------------------------------------
long ref_search(wchar_t *key, TRefCod2RefNo *a, long i, long step)
{
  if(step)
  {if( wcscmp(a[i].RefCode,key)<0) ref_search(key, a, i-step, step/2);
  else ref_search(key, a, i+step, step/2);
  }else
  return a[i].RefNo;
}
//---------------------------------------------------------------------------
__fastcall TRefEditForm::TRefEditForm(TComponent* Owner)
	: TForm(Owner)
{
}
//---------------------------------------------------------------------------
void __fastcall TRefEditForm::PrepareWord(void)
{
	try
	{
		WordApplication1->Connect();
	}
	catch(...)
	{
		ShowMessage("Unable to load Word");
	}
//	WordApplication1->Visible = true;
}
//---------------------------------------------------------------------------
void __fastcall TRefEditForm::Button_Click(TObject *Sender)
{

  wchar_t	*doc_contents;
  wchar_t	*ref_contents;
  wchar_t	*ref_codes;
  RangePtr	my_range;
  OleVariant	FileName;
  
  PrepareWord();

  if(OpenDialog1->Execute())
  { ref_fileName = OpenDialog1->FileName;
    FileName = ref_fileName;
    WordDocument1->ConnectTo(WordApplication1->Documents->Open(FileName));

    TablesPtr doc_tables;
              doc_tables=WordDocument1->Tables;
    Table* ref_table;
           ref_table=doc_tables->Item(1);
    Columns* ref_table_columns;
             ref_table_columns=ref_table->get_Columns();
    Column* ref_column;
            ref_column=ref_table_columns->get_First();
    Cells* cells;
           cells=ref_column->Cells;
    ref_count=cells->get_Count();

    Cell* cell;

    arr=new TRefCod2RefNo[ref_count];
    RangePtr  ref_range;
    for(long j=0;j<ref_count;j++)
                { cell=cells->Item(j+1);
                  ref_range=cell->get_Range();
                  ref_contents =ref_range->get_Text();
                  wcsncpy(arr[j].RefCode, ref_contents, 4);
                  arr[j].RefCode[4]='\0';
                  arr[j].RefNo=j+1;
                  arr[j].used=0;
                }
//    ref_sort(arr, ref_count);
    WordDocument1->Close();
    WordDocument1->Disconnect();
  }
/*
        ref_count=WordDocument1->Tables->Item(1)->get_Columns()->get_First()->Cells->get_Count();
        Cell* cell;
        arr=new TRefCod2RefNo[ref_count];
        for(long j=0;j<ref_count;j++)
        { cell=WordDocument1->Tables->Item(1)->get_Columns()->get_First()->Cells->Item(j+1);
          wcsncpy(arr[j].RefCode, cell->get_Range()->get_Text(), 4);
          arr[j].RefCode[4]='\0';
          arr[j].RefNo=j+1;
        }
         ListBox1->Items->Clear();
         ListBox2->Items->Clear();
         for(long j=0;j<ref_count;j++)
        {
         ListBox1->Items->Append((AnsiString)arr[j].RefCode);
         ListBox2->Items->Append(IntToStr(arr[j].RefNo));
        }
*/
}
//---------------------------------------------------------------------------
long TRefEditForm::RefCode2RefNo(wchar_t* ref_code)
{   long ind=0;
    for(long i=0;i<ref_count;i++)
       if(wcscmp(ref_code,arr[i].RefCode)==0)
         {ind=arr[i].RefNo; arr[i].used=1; break;}
    return ind;
}
//---------------------------------------------------------------------------
void __fastcall TRefEditForm::Button1Click(TObject *Sender)
{ 
  wchar_t	*doc_contents;
  wchar_t	*ref_contents;
  wchar_t	*ref_codes;
  RangePtr	my_range;

  AnsiString fileName;
  AnsiString refNo;

  for(int k=0;k<ListBox_Words->Items->Count;k++)
     { fileName=(AnsiString)ListBox_Words->Items->Strings[k];
       //-- We've got a file name - Open the document
       WordDocument1->ConnectTo(WordApplication1->Documents->Open((OleVariant)fileName));
//                /*
       ParagraphsPtr doc_pars;
                     doc_pars=WordDocument1->Paragraphs;
       long doc_pars_cnt=doc_pars->Count;

       Paragraph* doc_par;
       RangePtr par_range;
       wchar_t* par_contents;

       Words* par_words;
       RangePtr word_range;
       wchar_t* word_contents;

       for(long k=0;k<doc_pars_cnt;k++)
          {  doc_par=doc_pars->Item(k+1);
             par_range=doc_par->get_Range();
             par_words=par_range->get_Words();
             int par_word_cnt;
                 par_word_cnt=par_words->get_Count();
             for(int j=par_word_cnt;j>0;j--)
                { word_range=par_words->Item(j);
                  word_contents=word_range->get_Text();
                  int  d=wcslen(word_contents);
                  if(word_contents[0]=='#' || word_contents[1]=='#')
                    {
                     if(word_contents[0]=='#')
                     my_range=WordDocument1->Range((Variant)word_range->Start,(Variant)(word_range->Start+4));
                     if(word_contents[1]=='#')
                     my_range=WordDocument1->Range((Variant)(word_range->Start+1),(Variant)(word_range->Start+5));
//my_range->InsertBefore(StringToOleStr("\\cite{"));
//my_range->InsertAfter(StringToOleStr("}"));
                        ref_contents = my_range->get_Text();
                        long ref_n=RefCode2RefNo(ref_contents);
                        if(ref_n)
                           { refNo=IntToStr(ref_n);
//*  don't replace !!!*/                             my_range->set_Text(StringToOleStr(refNo));
                           }else
                           { // my_range->set_Text(StringToOleStr(""));
                              ListBox1->Items->Add("Reference "+(AnsiString)ref_contents+" is undefined, but used in "+fileName);
                           }
                    }
                }
  /*
               par_contents=par_range->get_Text();
               long d=wcslen(par_contents);
               for(long i=d; i>0;i--)
                  {
                    if(par_contents[i]=='#')
                      {
my_range = WordDocument1->Range((Variant)(par_range->Start+i), (Variant)(par_range->Start+i+4));
                        ref_contents = my_range->get_Text();
                        refNo=IntToStr(RefCode2RefNo(ref_contents));
// refNo=IntToStr(ref_search(ref_contents, arr, ref_count/2, ref_count/4));
                        my_range->set_Text(StringToOleStr(refNo));
                            }
                      }
  */

           }
  /*
  my_range = WordDocument1->Range(EmptyParam,EmptyParam);
  doc_contents = my_range->get_Text();
  long d=wcslen(doc_contents);
  for(long i=d-1; i>-1;i--)
                {
                  if(doc_contents[i]=='#')
                    { my_range = WordDocument1->Range((Variant) i, (Variant) i+4);

                      ref_contents = my_range->get_Text();

                      //my_range->set_Text(StringToOleStr(""));
//                      ListBox1->Items->Append((AnsiString)ref_contents);
                      refNo=IntToStr(RefCode2RefNo(ref_contents));
//                      ListBox2->Items->Append(refNo);
                      my_range->set_Text(StringToOleStr(refNo));
                      //my_range->InsertBefore(StringToOleStr("\\cite{"));
                      //my_range->InsertAfter(StringToOleStr("}"));
                    }
                }
*/
/*
      fileName=ExtractFilePath(fileName)+"ref_"+ExtractFileName(fileName);
      tagVARIANT* newFileName;
      newFileName=(tagVARIANT*) fileName.c_str();
      WordDocument1->SaveAs(newFileName);
      */

/*
TableOfContentsPtr  contents;
RangePtr	    contents_range;
contents=WordDocument1->get_TablesOfContents();
contents_range=contents->get_Range();
wchar_t	            *doc_toc;
doc_toc=contents_range->get_Text();
*/
      /*tagVARIANT* SaveChanges; SaveChanges=0;
      WordDocument1->Close(SaveChanges);*/
      WordDocument1->Save();
      WordDocument1->Close();
      WordDocument1->Disconnect();
      ListBox1->Items->Add(fileName+" proceed");
    }
    for(long i=0;i<ref_count;i++)
       if(arr[i].used==0)
         {ListBox1->Items->Add("Reference "+(AnsiString)arr[i].RefCode+" is never used");
         }
//    Application->NormalizeTopMosts();
}
//---------------------------------------------------------------------------
void __fastcall TRefEditForm::WMDropFiles(TWMDropFiles &message)
{
  AnsiString FileName;
  FileName.SetLength(MAX_PATH);

  int Count = DragQueryFile((HDROP)message.Drop, 0xFFFFFFFF, NULL, MAX_PATH);

  // index through the files and query the OS for each file name...
  for (int index = 0; index < Count; ++index)
  {
    // the following code gets the FileName of the dropped file.  I know it
    // looks cryptic but that's only because it is.  Hey, Why do you think
    // Delphi and C++ Builder are so popular anyway?
    FileName.SetLength(DragQueryFile((HDROP)message.Drop, index,
      FileName.c_str(), MAX_PATH));

    // examine the filename's extension.
    // If it's a Word file then ...
    if (UpperCase(ExtractFileExt(FileName)) == ".DOC")

    {
      ListBox_Words->Items->Add(FileName);
    }
  }
  // tell the OS that we're finished...
  DragFinish((HDROP) message.Drop);
}
//---------------------------------------------------------------------------
void __fastcall TRefEditForm::FormCreate(TObject *Sender)
{
  DragAcceptFiles(Handle, True);
}
//---------------------------------------------------------------------------
void __fastcall TRefEditForm::Button2Click(TObject *Sender)
{RangePtr tmp_range;
 AnsiString str;
 WordDocument2->ConnectTo(WordApplication1->Documents->Add());
 tmp_range=WordDocument2->get_Content();
 int cnt=ListBox1->Count;
 for(int i=0;i<cnt;i++)
 {str=(AnsiString)ListBox1->Items->Strings[i];
  tmp_range->InsertAfter(StringToOleStr(str));
 }
 tmp_range->InsertAfter(StringToOleStr("\n—œ»—Œ  ¬» Œ–»—“¿Õ»’ ƒ∆≈–≈À\n"));
 //-----
// wchar_t	*doc_contents;
 wchar_t	*ref_contents;
// RangePtr	my_range;
 OleVariant	FileName;
 FileName = ref_fileName;
 WordDocument1->ConnectTo(WordApplication1->Documents->Open(FileName));
 //-----
    TablesPtr doc_tables;
              doc_tables=WordDocument1->Tables;
    Table* ref_table;
           ref_table=doc_tables->Item(1);
    Columns* ref_table_columns;
             ref_table_columns=ref_table->get_Columns();
 //   Column* code_column;
 //           code_column=ref_table_columns->get_First();
 //   Cells* code_cells;
 //          code_cells=code_column->Cells;
 //   Cell* code_cell;
 //   RangePtr  code_range;

    Column* ref_column;
            ref_column=ref_table_columns->get_Last();
    Cells* ref_cells;
           ref_cells=ref_column->Cells;
    Cell* ref_cell;
    RangePtr  ref_range;
    long k=0;
    for(long j=0;j<ref_count;j++)
       {if(arr[j].used)
        {ref_cell=ref_cells->Item(j+1);
         ref_range=ref_cell->get_Range();
         ref_contents =ref_range->get_Text();
         tmp_range->InsertAfter(StringToOleStr(arr[j].RefCode));
         tmp_range->InsertAfter(StringToOleStr("\t"));
         k++;
         tmp_range->InsertAfter(StringToOleStr(IntToStr(k)));
         tmp_range->InsertAfter(StringToOleStr("\t"));
         tmp_range->InsertAfter(StringToOleStr(ref_contents));
        }
       }

 tmp_range->InsertAfter(StringToOleStr("\n—œ»—Œ  Õ≈¬» Œ–»—“¿Õ»’ ƒ∆≈–≈À\n"));
    for(long j=0;j<ref_count;j++)
       {if(arr[j].used==0)
        {ref_cell=ref_cells->Item(j+1);
         ref_range=ref_cell->get_Range();
         ref_contents =ref_range->get_Text();
         tmp_range->InsertAfter(StringToOleStr(arr[j].RefCode));
         tmp_range->InsertAfter(StringToOleStr("\t"));
         k++;
         tmp_range->InsertAfter(StringToOleStr(IntToStr(k)));
         tmp_range->InsertAfter(StringToOleStr("\t"));
         tmp_range->InsertAfter(StringToOleStr(ref_contents));
        }
       }
 WordDocument1->Close();
 WordDocument1->Disconnect();
 WordDocument2->Save();
 WordDocument2->Close();
 WordDocument2->Disconnect();

}
//---------------------------------------------------------------------------

void __fastcall TRefEditForm::Button3Click(TObject *Sender)
{
  wchar_t	*doc_contents;
  wchar_t	*ref_contents;
  wchar_t	*ref_codes;
  RangePtr	my_range;

  AnsiString fileName;
  AnsiString refNo;

  ListBox1->Items->Clear();

  for(int k=0;k<ListBox_Words->Items->Count;k++)
     { fileName=(AnsiString)ListBox_Words->Items->Strings[k];
       //-- We've got a file name - Open the document
       WordDocument1->ConnectTo(WordApplication1->Documents->Open((OleVariant)fileName));

       ParagraphsPtr doc_pars;
                     doc_pars=WordDocument1->Paragraphs;
       long doc_pars_cnt=doc_pars->Count;

       Paragraph* doc_par;
       RangePtr par_range;
       wchar_t* par_contents;

       Words* par_words;
       RangePtr word_range;
       wchar_t* word_contents;

       for(long k=0;k<doc_pars_cnt;k++)
          {  doc_par=doc_pars->Item(k+1);
             par_range=doc_par->get_Range();
             par_words=par_range->get_Words();
             int par_word_cnt;
                 par_word_cnt=par_words->get_Count();
/**/
             unsigned int ref_open=0;
             unsigned int brk_open=0;
             wchar_t* refs_bracked[25];
             long refs_ind[25];
             int refs_cnt=0;

             RangePtr brackets_beg_range;
             RangePtr brackets_end_range;
/**/
             for(int j=par_word_cnt;j>0;j--)
                { word_range=par_words->Item(j);
                  word_contents=word_range->get_Text();

                  if(word_contents[0]=='#' || word_contents[1]=='#')
                    {
                     if(word_contents[0]=='#')
                     {my_range=WordDocument1->Range((Variant)word_range->Start,(Variant)(word_range->Start+4));
                      if(!brk_open)
                        {brackets_end_range=WordDocument1->Range((Variant)word_range->Start,(Variant)(word_range->Start+4));
                         brk_open=1;
                        }

                     }
                     if(word_contents[1]=='#')
                     {my_range=WordDocument1->Range((Variant)(word_range->Start+1),(Variant)(word_range->Start+5));
                      ref_open=1;
                      if(!brk_open)
                        {brackets_end_range=WordDocument1->Range((Variant)word_range->Start+1,(Variant)(word_range->Start+5));
                         brk_open=1;
                        }
                      brk_open=0;
                     }
/**/
                     refs_bracked[refs_cnt]=my_range->get_Text();
                     refs_cnt++;

                     if(ref_open)
                     {
                         for(int i=0;i<refs_cnt;i++)
                            {
                             refs_ind[i]=RefCode2RefNo(refs_bracked[i]);
                            }
                         for(int j=0;j<refs_cnt;j++)
                            for(int k=0;k<refs_cnt;k++)
                            {if(refs_ind[j]<refs_ind[k])
                               {long tmp=refs_ind[j];
                                refs_ind[j]=refs_ind[k];
                                refs_ind[k]=tmp;
                               }
                            }
                         AnsiString a;
                         AnsiString b;
                         for(int i=0;i<refs_cnt-1;i++)
                            {b=IntToStr(refs_ind[i]);
                             b+=", ";
                             a+=b;
                            }
                         b=IntToStr(refs_ind[refs_cnt-1]);
                         a+=b;
//my_range=WordDocument1->Range((Variant)(word_range->Start+1),(Variant)(word_range->Start+refs_cnt*6-1));
my_range=WordDocument1->Range((Variant)(word_range->Start+1),(Variant)(brackets_end_range->Start+4));
my_range->set_Text(StringToOleStr(a));
                         refs_cnt=0;
                         ref_open=0;
                     }
                    }
                }
           }
     WordDocument1->Save();
     WordDocument1->Close();
     WordDocument1->Disconnect();
     ListBox1->Items->Add(fileName+" proceed");
    }
}
//---------------------------------------------------------------------------

void __fastcall TRefEditForm::Button4Click(TObject *Sender)
{
 Application->Terminate();        
}
//---------------------------------------------------------------------------

