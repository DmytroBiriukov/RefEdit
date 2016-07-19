using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Data;
using MySql.Data;
using MySql.Data.MySqlClient;
using Microsoft.Office.Tools.Ribbon;
//using Microsoft.Office.Tools.Word;
using Microsoft.Office.Interop.Word;
using System.Reflection;
using System.Windows.Forms;
using System.Data.OleDb;


namespace RefEditor
{   
    public partial class Ribbon1
    {   /* to manage mySQL database
         */ 
        private const String myConnectionString = @"server=mysql301.1gb.ua;port=3306;database=gbua_x_datac31a;userid=gbua_x_datac31a;password=a6b0a10bem1;";
        private MySqlConnection con;
        private MySqlDataReader reader;
        private const uint idOwner = 1;
        /* to manage local database (accdb file)
        */
        private OleDbConnection connection;
        private OleDbCommand command;
        private OleDbDataAdapter adapter;
        private DataSet dataset;


        const int symbols_num = 36;
        private char[] symbolesInCode = new char[symbols_num] {'0', '1', '2', '3', '4', '5', '6', '7', '8', '9', 'A', 'B', 'C', 'D', 'E', 'F', 
                                                   'G', 'H', 'J', 'I', 'K', 'L', 'M', 'N', 'O', 'P', 'Q', 'R', 'S', 'T', 'U', 'V', 
                                                   'W', 'X', 'Y', 'Z'};
        /*, 'a', 'b', 'c', 'd', 'e', 'f', 'g', 'h', 'j', 'i', 'k', 'l', 
        'm', 'n', 'o', 'p', 'q', 'r', 's', 't', 'u', 'v', 'w', 'x', 'y', 'z', '$', '&'};
         */ 


        Range rng;

        List<TRefCod2RefNo> arr = new List<TRefCod2RefNo>();
        List<TRefCod2RefNo> hyperlinks = new List<TRefCod2RefNo>();
        uint ref_count;
        public long RefCode2RefNo(string ref_code)
        {
            TRefCod2RefNo result = new TRefCod2RefNo();
            result.RefNo = 0;
            result = arr.Find(delegate(TRefCod2RefNo refc) { return refc.RefCode == ref_code; });
            if (result != null) return result.RefNo; else return 0;

        }
        public void ref_sort()
        {
            arr.Sort();
        }
        public long ref_search(string key)
        {
            TRefCod2RefNo result = new TRefCod2RefNo();
            result.RefNo = 0;
            result = arr.Find(delegate(TRefCod2RefNo refc) { return refc.RefCode == key; });
            if (result != null) return result.RefNo; else return 0;
        }
        private static bool FindRef(TRefCod2RefNo refc, string ref_code)
        {
            if (refc.RefCode == ref_code) return true; else return false;
        }

        private void Ribbon1_Load(object sender, RibbonUIEventArgs e)
        {

        }

        private void button1_Click(object sender, RibbonControlEventArgs e)
        {
            object start = 0;
            object end = 0;
            Microsoft.Office.Interop.Word.Range rng = Globals.ThisAddIn.Application.ActiveDocument.Range(ref start, ref end);
            rng.Select();
            Microsoft.Office.Interop.Word.Range current_rng = rng;

            // http://forum.codecall.net/topic/71422-connecting-to-a-mysql-database-in-c/
 
            con = null;
            reader = null;
            try
            {
                con = new MySqlConnection(myConnectionString);
                con.Open(); //open the connection
                /*
                //This is the mysql command that we will query into the db.
                //It uses Prepared statements and the Placeholder is @name.
                //Using prepared statements is faster and secure.
                String cmdText = "INSERT INTO myTable(name) VALUES(@name)";
                MySqlCommand cmd = new MySqlCommand(cmdText, con);
                cmd.Prepare();
                //we will bound a value to the placeholder
                cmd.Parameters.AddWithValue("@name", "your value here");
                cmd.ExecuteNonQuery(); //execute the mysql command
                */

                //We will need to SELECT all or some columns in the table via this command
                String cmdText = "SELECT * FROM merg_person";
                MySqlCommand cmd = new MySqlCommand(cmdText,con);
                reader = cmd.ExecuteReader(); //execure the reader
                /*The Read() method points to the next record It return false if there are no more records else returns true.*/
                while (reader.Read())
                {            /*reader.GetString(0) will get the value of the first column of the table myTable because we selected all columns using SELECT * (all); the first loop of the while loop is the first row; the next loop will be the second row and so on...*/
	                //Console.WriteLine(reader.GetString(0));
                    current_rng.Text = reader.GetString(1);
                    current_rng = Globals.ThisAddIn.Application.ActiveDocument.Range(ref start, ref end);

                }
            }
            catch (MySqlException err)
            {
                Console.WriteLine("Error: " + err.ToString());
            }
            finally
            {
                if (con != null)
                {
                    con.Close(); //close the connection
                }
            } //remember to close the connection after accessing the database

        }

        private void button2_Click(object sender, RibbonControlEventArgs e)
        {   Form1 frm = new Form1();
            frm.Show();
        }

        private void button3_Click(object sender, RibbonControlEventArgs e)
        {   openFileDialog1.ShowDialog();
        }

        private void openFileDialog1_FileOk(object sender, System.ComponentModel.CancelEventArgs e)
        {  initializeReferences(openFileDialog1.FileName);
        }

        private void initializeReferences(object filename)
        {
            object missing = Type.Missing;
            Microsoft.Office.Interop.Word.Document doc = Globals.ThisAddIn.Application.Documents.Open(ref filename, ref missing, ref missing, ref missing, ref missing,
                ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing);
            Microsoft.Office.Interop.Word.Table nowTable = doc.Tables[1];
            for (int rowPos = 1; rowPos <= nowTable.Rows.Count; rowPos++)
            { //for (int columPos = 1; columPos <= nowTable.Columns.Count; columPos++)
                String strContent;
                TRefCod2RefNo n_item = new TRefCod2RefNo();
                n_item.RefNo = rowPos;
                strContent = nowTable.Cell(rowPos, 1).Range.Text.ToString();
                n_item.RefCode = strContent.Substring(0, 4);
                strContent = nowTable.Cell(rowPos, 3).Range.Text.ToString();
                n_item.Ref_entity = strContent.Trim();
                arr.Add(n_item);
            }

            object template = Missing.Value; //No template.
            object newTemplate = Missing.Value; //Not creating a template.
            object documentType = Missing.Value; //Plain old text document.
            object visible = true;  //Show the doc while we work.
            Microsoft.Office.Interop.Word.Document doc_new = Globals.ThisAddIn.Application.Documents.Add(ref template, ref newTemplate, ref documentType, ref visible);
            //ref_sort();
            foreach (TRefCod2RefNo a_item in arr)
            {
                rng = doc_new.Range(0, 0);
                rng.Text = a_item.Ref_entity + " <" + a_item.RefCode + ">";

            }

        }

        private void button4_Click(object sender, RibbonControlEventArgs e)
        {   initializeReferences("D://Desktop//Doc1.docx");           
        }

        private void button5_Click(object sender, RibbonControlEventArgs e)
        {   Form2 frm = new Form2();
            frm.Show();
        }

        private void operate_document(Document doc)
        {   Microsoft.Office.Interop.Word.Hyperlinks myLinks = doc.Hyperlinks;
            rng = doc.Range(0, 0);
            rng.Select();
            Microsoft.Office.Interop.Word.Selection currentSelection = Globals.ThisAddIn.Application.Selection;
            for (int i = 1; i <= myLinks.Count; i++)
            {
                TRefCod2RefNo result = new TRefCod2RefNo();
                object index = (object)i;
                Microsoft.Office.Interop.Word.Hyperlink link = myLinks.get_Item(ref index);
                //               myLinks.get_Item(ref index).ScreenTip = myLinks.get_Item(ref index).ScreenTip.ToUpper();

                String key = "#" + link.Address.ToString();

                rng.InsertBefore(key);

                result = arr.Find(delegate(TRefCod2RefNo refc) { return refc.RefCode == key; });

                //result.RefCode=arr.Find(delegate(TRefCod2RefNo refc){ return refc.RefCode == key;}).RefCode;

                if (result != null)
                {  //currentSelection.TypeText(myLinks[i].TextToDisplay.ToString());
                    //currentSelection.TypeParagraph();
                    //   link.ScreenTip.ToUpper();// = result.Ref_entity;
                    myLinks.get_Item(ref index).ScreenTip = result.Ref_entity;
                    myLinks.get_Item(ref index).ScreenTip.ToUpper();
                    rng.InsertBefore(" found ");
                }


            }

            /*
             * 
             *             try
            {
                conn.Open();
              
                System.Data.DataTable table = conn.GetSchema("merg_institution");
                foreach (System.Data.DataRow row in table.Rows)
                {
                    foreach (System.Data.DataColumn col in table.Columns)
                    {
                        rng.InsertAfter(col.ColumnName.ToString() + row[col].ToString());
                    }

                }
               
                conn.Close();
            }
            catch (Exception ex)
            {
                rng.InsertAfter(ex.ToString());
            }
             */


            //doc.Save();
        }

        private void button6_Click(object sender, RibbonControlEventArgs e)
        {   Document doc = Globals.ThisAddIn.Application.ActiveDocument;
            operate_document(doc);
        }

        private void button7_Click(object sender, RibbonControlEventArgs e)
        {
            Document doc = Globals.ThisAddIn.Application.ActiveDocument;
            Range rng = doc.Range(0, 0);
            MySqlConnection conn = new MySqlConnection(myConnectionString);
            try
            {
                conn.Open();
                /*
                System.Data.DataTable table = conn.GetSchema("merg_institution");
                foreach (System.Data.DataRow row in table.Rows)
                {
                    foreach (System.Data.DataColumn col in table.Columns)
                    {
                        rng.InsertAfter(col.ColumnName.ToString() + row[col].ToString());
                    }

                }
                 */
                conn.Close();
            }
            catch (Exception ex)
            {
                rng.InsertAfter(ex.ToString());
            }
         
        }

        private void button7_Click_1(object sender, RibbonControlEventArgs e)
        {
            var app = Globals.ThisAddIn.Application;
            foreach (var item in app.Documents)
            {   var doc = (Document)item;
                operate_document(doc);
            }
        }

        private void button8_Click(object sender, RibbonControlEventArgs e)
        {
            var app = Globals.ThisAddIn.Application;
            foreach (var item in app.Documents)
            {   var doc = (Document)item;               
                // Collect references
                Microsoft.Office.Interop.Word.Hyperlinks myLinks = doc.Hyperlinks;
                for (int i = 1; i <= myLinks.Count; i++)
                {   object index = (object)i;
                    Microsoft.Office.Interop.Word.Hyperlink link = myLinks.get_Item(ref index);
                    TRefCod2RefNo n_item = new TRefCod2RefNo();
                    //n_item.RefNo = i;
                    n_item.RefNo = 0;
                    n_item.used = 0;
                    n_item.RefCode = link.SubAddress.ToString();
                    if (link.ScreenTip != null) n_item.Ref_entity = link.ScreenTip.ToString();
                    else n_item.Ref_entity = "";
                    hyperlinks.Add(n_item);
                }
                // [end] Collect references
            }

            object template = Missing.Value; 
            object newTemplate = Missing.Value; 
            object documentType = Missing.Value; 
            object visible = true; 
            Document new_doc = app.Documents.Add(ref template, ref newTemplate, ref documentType, ref visible);

            List<TRefCod2RefNo> refs = hyperlinks.OrderBy(o => o.Ref_entity).ToList();
            
            Int32 ind = 0;
            while (ind < refs.Count - 1)
            { if (String.Equals(refs[ind].Ref_entity, refs[ind + 1].Ref_entity, StringComparison.OrdinalIgnoreCase) || String.Equals(refs[ind].RefCode, refs[ind + 1].RefCode, StringComparison.OrdinalIgnoreCase)) refs.RemoveAt(ind);
              else ind++;
            }

            for (ind = 0; ind < refs.Count; ind++ )
            {
                refs[ind].RefNo = (ind + 1);
            }

            //List<TRefCod2RefNo> refs = hyperlinks.OrderBy(o => o.RefCode).Distinct().ToList();
            new_doc.Words.First.InsertAfter("Список використаних джерел");

            object start = 0; // new_doc.Content.End; //new_doc.Content.Start;
            object end = 0; // new_doc.Content.End;
            Range rng = new_doc.Range(ref start, ref end);
            
            // Add the table.
            Object defaultTableBehavior = Type.Missing;
            Object autoFitBehavior = Type.Missing;
            new_doc.Tables.Add(rng, 1, 3, ref defaultTableBehavior, ref autoFitBehavior);

            Table tbl = new_doc.Tables[1];
            tbl.Range.Font.Size = 8;
            tbl.Range.Font.Name = "Verdana";

            Object style = "Table Grid 8";
            tbl.set_Style(ref style);
            tbl.ApplyStyleFirstColumn = false;
            tbl.ApplyStyleLastColumn = false;
            tbl.ApplyStyleLastRow = false;

            // Insert header text and format the columns.
            tbl.Cell(1, 1).Range.Text = "Code";

            Range rngCell;
            rngCell = tbl.Cell(1, 2).Range;
            rngCell.Text = "No.";
            rngCell.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphCenter;
            rngCell = tbl.Cell(1, 3).Range;
            rngCell.Text = "Title";
            rngCell.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphLeft;

            Object beforeRow = Type.Missing;
            
            ind = 2;

            foreach (TRefCod2RefNo a_item in refs)
            {
                tbl.Rows.Add(ref beforeRow);
                tbl.Cell(ind, 1).Range.Text = a_item.RefCode.ToString();
                tbl.Cell(ind, 2).Range.Text = a_item.RefNo.ToString();
                  //string.Format("{0:N0}", fi.Length);
                tbl.Cell(ind, 3).Range.Text = a_item.Ref_entity.ToString();
                  //  string.Format("{0:g}", fi.LastWriteTime);
                ind++;
                //new_doc.Words.Last.InsertAfter(a_item.RefNo.ToString()+" - "+ a_item.RefCode.ToString() + " - " + a_item.Ref_entity.ToString());
            }


            foreach (var item in app.Documents)
            {
                var doc = (Document)item;
                // Set references
                Microsoft.Office.Interop.Word.Hyperlinks myLinks = doc.Hyperlinks;
                for (int i = 1; i <= myLinks.Count; i++)
                {
                    object index = (object)i;
                    String key = myLinks.get_Item(ref index).Address;
                    TRefCod2RefNo result = new TRefCod2RefNo();
                    
                    result.RefNo = 0;
                    result = refs.Find(delegate(TRefCod2RefNo refc) { return refc.RefCode == key; });
                    if (result != null) myLinks.get_Item(ref index).TextToDisplay = result.RefNo.ToString(); else myLinks.get_Item(ref index).TextToDisplay = "nf";

                }
                // [end] Set references
            }
            
            //Save the file, use default values except for filename.
            /*
            object fileName = Environment.CurrentDirectory + "\\example2_new";
            object optional = Missing.Value; 
            */
#if OFFICEXP
          //doc.SaveAs2000( ref fileName,
#else
            //new_doc.SaveAs(ref fileName,
#endif
 //ref optional, ref optional, ref optional,ref optional, ref optional,ref optional,ref optional, ref optional, ref optional, ref optional);
            /*
            // Now use the Quit method to cleanup.
            object saveChanges = true;
            app.Quit(ref saveChanges, ref optional, ref optional);
            */
        }

        private void button9_Click(object sender, RibbonControlEventArgs e)
        {/*  This function finds all fragments in the text braced with [ ] - which corresponds for references
          * and replace this franments by @@@
          * -
          * will be used to edit references due to formal requirements (e.g., "[1,2, 3-5]" ) 
          */
            var app = Globals.ThisAddIn.Application;
            Range range = app.ActiveDocument.Content;
            
            object missing = System.Reflection.Missing.Value;
            object findtext = "[";
            object f = false;
            object findreplacement = "@";
            object findforward = false;
            object findformat = true;
            object findwrap = WdFindWrap.wdFindContinue;
            object findmatchcase = false;
            object findmatchwholeword = false;
            object findmatchwildcards = false;
            object findmatchsoundslike = false;
            object findmatchallwordforms = false;
            object findreplace = WdReplace.wdReplaceNone; //WdReplace.wdReplaceAll;

            range.Find.Execute(findtext,findmatchcase,findmatchwholeword, findmatchwildcards, findmatchsoundslike, findmatchallwordforms, findforward,findwrap, findformat, findreplacement, findreplace, missing, missing,  missing, missing);
            while (range.Find.Found)
            {
                object start = range.Start; //app.ActiveDocument.Content.Start;
                object end = app.ActiveDocument.Content.End;
                
                Range end_range = app.ActiveDocument.Range(ref start, ref end);
                findtext = "]";
                //end_range
                end_range.Find.Execute(findtext, findmatchcase, findmatchwholeword, findmatchwildcards, findmatchsoundslike, findmatchallwordforms, findforward, findwrap, findformat, findreplacement, findreplace, missing, missing, missing, missing);
                if (end_range.Find.Found)
                {
                    Range found_range = app.ActiveDocument.Range(range.Start, end_range.End);
                    //found_range.Delete();
                    found_range.Text = "@@@";
                }
                findtext = "[";
                range.Find.Execute(findtext, findmatchcase, findmatchwholeword, findmatchwildcards, findmatchsoundslike, findmatchallwordforms, findforward, findwrap, findformat, findreplacement, findreplace, missing, missing, missing, missing);
            }


        }

        private void button10_Click(object sender, RibbonControlEventArgs e)
        {
            var app = Globals.ThisAddIn.Application;
            Document doc = app.ActiveDocument;
            Range rng;
            Microsoft.Office.Interop.Word.Hyperlinks allLinks = doc.Hyperlinks;
            foreach (Footnote fnote in doc.Footnotes )
            {   string test_file_Path = "#";
                object linkAddr = test_file_Path;
                string test_bookmark= "b2r";
                object linkSubAddr = test_bookmark;
                string screenTip = fnote.Range.Text; 
                object linkScreenTip = screenTip;
                string test_todisplay = test_file_Path + test_bookmark;
                object linkToDisplay = test_todisplay;
                rng = fnote.Reference;                
                //fnote.Delete();
                Microsoft.Office.Interop.Word.Hyperlink hl = allLinks.Add(rng, ref linkAddr, ref linkSubAddr, ref linkScreenTip, ref linkToDisplay);
                hl.Range.InsertBefore(" [");
                hl.Range.InsertAfter("]");

                
            }

        }

        private void button11_Click(object sender, RibbonControlEventArgs e)
        {   Form3 frm = new Form3();
            frm.Show();    
        }

        /*
         * exapmle of text formating in the document 
         */
        private void RangeFormat()
        {   var app = Globals.ThisAddIn.Application;
            Document doc = app.ActiveDocument;
            // Set the Range to the first paragraph. 
            Range rng = doc.Paragraphs[1].Range;

            // Change the formatting. To change the font size for a right-to-left language,  
            // such as Arabic or Hebrew, use the Font.SizeBi property instead of Font.Size.
            rng.Font.Size = 14;
            rng.Font.Name = "Arial";
            rng.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphCenter;

            rng.Select();
            MessageBox.Show("Formatted Range");

            // Undo the three previous actions.  
            object numTimes3 = 3;
            doc.Undo(ref numTimes3);

            rng.Select();
            MessageBox.Show("Undo 3 actions");

            // Apply the Normal Indent style.  
            object indentStyle = "Normal Indent";
            rng.set_Style(ref indentStyle);

            rng.Select();
            MessageBox.Show("Normal Indent style applied");

            // Undo a single action.  
            object numTimes1 = 1;
            doc.Undo(ref numTimes1);

            rng.Select();
            MessageBox.Show("Undo 1 action");
        }

        private void button12_Click(object sender, RibbonControlEventArgs e)
        {
            var app = Globals.ThisAddIn.Application;
            Document doc = app.ActiveDocument;
            Range rng;
            List<string> footnote_list=new List<string>();
            foreach (Footnote fnote in doc.Footnotes) footnote_list.Add(fnote.Range.Text);
            footnote_list.Sort();
            List<string> footnote_list_unique=footnote_list.Distinct().ToList();

            object template = Missing.Value;
            object newTemplate = Missing.Value;
            object documentType = Missing.Value;
            object visible = true;
            Document new_doc = app.Documents.Add(ref template, ref newTemplate, ref documentType, ref visible);
 
            new_doc.Words.First.InsertAfter("Список використаних джерел");

            object start = 0; // new_doc.Content.End; //new_doc.Content.Start;
            object end = 0; // new_doc.Content.End;
            rng = new_doc.Range(ref start, ref end);

            // Add the table.
            Object defaultTableBehavior = Type.Missing;
            Object autoFitBehavior = Type.Missing;
            new_doc.Tables.Add(rng, 1, 3, ref defaultTableBehavior, ref autoFitBehavior);

            Table tbl = new_doc.Tables[1];
            
            // Insert header text and format the columns.
            tbl.Cell(1, 1).Range.Text = "Code";

            Range rngCell;
            rngCell = tbl.Cell(1, 2).Range;
            rngCell.Text = "No.";
            rngCell = tbl.Cell(1, 3).Range;
            rngCell.Text = "Title";
            
            Object beforeRow = Type.Missing;

            Int32 ind = 2;

            int[] codeArray = new int[3];
            codeArray[0] = 2;
            codeArray[1] = 6;
            codeArray[2] = 11;

            foreach (string a_item in footnote_list_unique)
            {
                tbl.Rows.Add(ref beforeRow);
                if (codeArray[0] > (symbols_num-1)) 
                { codeArray[0] = 0; 
                  codeArray[1] ++;
                  if (codeArray[1] > (symbols_num - 1)) 
                  { codeArray[1] = 0; 
                    codeArray[2] ++;                     
                  }
                }

                tbl.Cell(ind, 1).Range.Text = "#" + symbolesInCode[codeArray[2]].ToString()+ symbolesInCode[codeArray[1]].ToString()+ symbolesInCode[codeArray[0]].ToString();
                tbl.Cell(ind, 2).Range.Text = "";
                tbl.Cell(ind, 3).Range.Text = a_item.ToString();
                codeArray[0]++;
                ind++;
                
            }







        }

        private void button13_Click(object sender, RibbonControlEventArgs e)
        {
            connection = new OleDbConnection();
            command = new OleDbCommand();
            adapter = new OleDbDataAdapter();
            dataset = new DataSet();
            // edit database location using  AppDomain.CurrentDomain.BaseDirectory   of similar constants
            connection.ConnectionString = @"Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + AppDomain.CurrentDomain.BaseDirectory + "/refs.accdb;" +
            "Persist Security Info=False";
            command.Connection = connection;
            command.CommandText = "SELECT * FROM reference";
            adapter.SelectCommand = command;
            List<TRefCod2RefNo> local_refs = new List<TRefCod2RefNo>();
            var app = Globals.ThisAddIn.Application;
            try
            {
                adapter.Fill(dataset, "reference");
                
                
                
                DataRowCollection rows = dataset.Tables["reference"].Rows;
                
                foreach (DataRow row in rows)
                {
                    TRefCod2RefNo item = new TRefCod2RefNo();
                    item.RefCode= row[1].ToString();
                    item.Ref_entity = row[2].ToString();
                    local_refs.Add(item);
                }


                //print all RefCod2RefNo to new word document
                /*
                object template = Missing.Value; //No template.
                object newTemplate = Missing.Value; //Not creating a template.
                object documentType = Missing.Value; //Plain old text document.
                object visible = true;  //Show the doc while we work.
                Microsoft.Office.Interop.Word.Document doc_new = Globals.ThisAddIn.Application.Documents.Add(ref template, ref newTemplate, ref documentType, ref visible);
                rng = doc_new.Range(0, 0);
                rng.InsertAfter("Rows number: "+rows.Count.ToString());
                foreach (TRefCod2RefNo a_item in local_refs) rng.InsertAfter(" <" + a_item.RefCode + "> <" + a_item.Ref_entity + ">");
                */
                
                // replace all footnotes by appropriate hyperlinks
                
                
                Document doc = app.ActiveDocument;                
                Microsoft.Office.Interop.Word.Hyperlinks allLinks = doc.Hyperlinks;
                Range rng;

                int[] codeArray = new int[3];
                // #BNE
                codeArray[0] = 14;
                codeArray[1] = 23;
                codeArray[2] = 11;

                foreach (Footnote fnote in doc.Footnotes)
                {
                    TRefCod2RefNo result = new TRefCod2RefNo();
                    result = local_refs.Find(delegate(TRefCod2RefNo refc) { return ( String.Compare(refc.Ref_entity, fnote.Range.Text, true)==0 ); });
                    
                    string test_file_Path = "";
                    object linkAddr = test_file_Path;
                    string test_bookmark = "b2r";
                    string screenTip = "";
                    if (result != null)
                    {
                        test_bookmark = result.RefCode;
                        screenTip = result.Ref_entity;
                    }
                    else
                    {
                        codeArray[0]++;
                        if (codeArray[0] > (symbols_num - 1))
                        {
                            codeArray[0] = 0;
                            codeArray[1]++;
                            if (codeArray[1] > (symbols_num - 1))
                            {
                                codeArray[1] = 0;
                                codeArray[2]++;
                            }
                        }

                        test_bookmark = "#" + symbolesInCode[codeArray[2]].ToString() + symbolesInCode[codeArray[1]].ToString() + symbolesInCode[codeArray[0]].ToString();
                        screenTip = fnote.Range.Text; //fnote.Reference.Text.ToString(); 
                        //rows.Add();
                        TRefCod2RefNo n_item = new TRefCod2RefNo();
                        n_item.RefCode=test_bookmark;
                        n_item.Ref_entity = screenTip;
                        local_refs.Add(n_item);


                        adapter.InsertCommand = new OleDbCommand("INSERT INTO reference (code, title) VALUES (@code , @title)", connection);
                        adapter.InsertCommand.Parameters.Add("code", OleDbType.VarChar, 4, test_bookmark);
                        adapter.InsertCommand.Parameters.Add("title", OleDbType.VarChar, 1023, screenTip);
                        adapter.Update(dataset, "reference");
                        /*

                            connection.Open();

                            // Create an OleDbDataAdapter and provide it with an INSERT command.
                            var adapter = new OleDbDataAdapter();
                            
                        }
                    */


                    }
                    object linkSubAddr = test_bookmark;                    
                    object linkScreenTip = screenTip;
                    string test_todisplay = test_bookmark;
                    object linkToDisplay = test_todisplay;
                    rng = fnote.Reference;
                    Microsoft.Office.Interop.Word.Hyperlink hl = allLinks.Add(rng, ref linkAddr, ref linkSubAddr, ref linkScreenTip, ref linkToDisplay);
                    hl.Range.InsertBefore(" [");
                    hl.Range.InsertAfter("]");


                }
                /**/


            }
            catch (OleDbException)
            {
                MessageBox.Show("Error occured while connecting to database.");
            }


            object template = Missing.Value;
            object newTemplate = Missing.Value;
            object documentType = Missing.Value;
            object visible = true;
            
            Document new_doc = app.Documents.Add(ref template, ref newTemplate, ref documentType, ref visible);

            object start = 0; // new_doc.Content.End; //new_doc.Content.Start;
            object end = 0; // new_doc.Content.End;
            Range rng1 = new_doc.Range(ref start, ref end);

            rng1.InsertAfter("Список використаних джерел");


            // Add the table.
            Object defaultTableBehavior = Type.Missing;
            Object autoFitBehavior = Type.Missing;
            new_doc.Tables.Add(rng1, 1, 3, ref defaultTableBehavior, ref autoFitBehavior);

            Table tbl = new_doc.Tables[1];

            // Insert header text and format the columns.
            tbl.Cell(1, 1).Range.Text = "Code";
            Range rngCell;
            rngCell = tbl.Cell(1, 2).Range;
            rngCell.Text = "No.";
            rngCell = tbl.Cell(1, 3).Range;
            rngCell.Text = "Title";
            Object beforeRow = Type.Missing;
            Int32 ind = 2;
            foreach ( TRefCod2RefNo item in local_refs)
            {
                tbl.Rows.Add(ref beforeRow);
                tbl.Cell(ind, 1).Range.Text = item.RefCode;
                tbl.Cell(ind, 2).Range.Text = "";
                tbl.Cell(ind, 3).Range.Text = item.Ref_entity;
                ind++;
            }
        }

        private void button14_Click(object sender, RibbonControlEventArgs e)
        {   
            var app = Globals.ThisAddIn.Application;
            
            foreach (var item in app.Documents)
            {   int i0 = 0; 
                int i1 = 0;
                int i2 = 0;
                var doc = (Document)item;
                // Collect references
                Microsoft.Office.Interop.Word.Hyperlinks myLinks = doc.Hyperlinks;
                for (int i = 1; i <= myLinks.Count; i++)
                {   object index = (object)i;
                    Microsoft.Office.Interop.Word.Hyperlink link = myLinks.get_Item(ref index);
                    string key = "#";
                    if ((link.SubAddress.Length == 4) && String.Compare(link.Address.ToString(), key, true) == 0) i1++;

                    if (link.SubAddress != null && link.Address != null)
                    {
                       // link.SubAddress = link.Address;
                       // link.Address = key;
                        if ( (link.SubAddress.Length == 4) &&  String.Compare(link.Address.ToString(), key, true)==0 ) i1++;
                        else i2++;
                    }
                    else i0++;
                    

                    /*
                    if (link.SubAddress != null)
                    {
                        if (link.SubAddress.Length > 1)
                        {
                            link.Address = link.SubAddress;
                            link.SubAddress = null;

                            if (link.Address != null)
                            {
                                if (link.Address.Length == 4) i0++;
                                if (link.Address.Length > 4) i1++;
                                if (link.Address.Length < 4) i2++;
                            }
                        }
                        else
                        {
                            link.SubAddress = null;
                            if (link.Address != null)
                            {
                                if (link.Address.Length == 4) i0++;
                                if (link.Address.Length > 4) i1++; 
                                if (link.Address.Length < 4) i2++; 
                            }
                        }


                        // string subadd = link.SubAddress.ToString();
                        // if (subadd.Contains("#")) myLinks.get_Item(ref index).SubAddress.Substring(1);
                                                  
                    }
                */
                    
                }
                MessageBox.Show("The calculations are complete linksCount=" + myLinks.Count.ToString() + "  empty ref =" + i0.ToString() + " right =" + i1.ToString() + " wrong=" + i2.ToString(), "My Application", MessageBoxButtons.OKCancel, MessageBoxIcon.Asterisk);
            }
            
        }

    }

    public class TRefCod2RefNo : IComparable<TRefCod2RefNo>
    {
        public long RefNo;
        public string RefCode;
        public String Ref_entity;
        public uint used;
        public int CompareTo(TRefCod2RefNo value)
        {
            return String.Compare(value.RefCode,this.RefCode);
           // return String.Compare(value.Ref_entity, this.Ref_entity);
        }
    }
}
